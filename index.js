import puppeteer from 'puppeteer';
import Excel from 'exceljs'


(async () => {
  const modoBusqueda = true
  const QUERY_NEGOCIOS = '.Nv2PK'
  const waitTimeout = 350000
  const outputFile = 'Output.xlsx'
  const backupFile = 'Output-Backup.xlsx'
  const browser = await puppeteer.launch({headless: false});
  const page = (await browser.pages())[0];

  const output = new Excel.Workbook()
  await output.xlsx.readFile(outputFile)
  const worksheet = output.getWorksheet("Hoja1");
  worksheet.columns = [
    {header: 'link', key: 'link'},
    {header: 'name', key: 'name'},
    {header: 'web', key: 'web'},
    {header: 'email', key: 'email'},
    {header: 'phone', key: 'phone'},
    {header: 'facebook', key: 'facebook'},
    {header: 'instagram', key: 'instagram'},
    {header: 'linkedin', key: 'linkedin'},
    {header: 'query', key: 'query'}
  ]
  const collected = worksheet.getColumn('link').values.slice(2) || []

  const input = new Excel.Workbook()
  await input.xlsx.readFile('Input.xlsx')
  const ws = input.getWorksheet("Hoja1")
  ws.columns[{header: 'links', key: 'links'}]
  const extract = ws.getColumn(1).values.slice(2)  

  let negocios
  let extraidos = collected.length
  if(modoBusqueda){
    const queries = Array.from(new Set(worksheet.getColumn('query').values))
    queries.pop()
    for(let query of extract){
      const search = query.text || query.result || query
      if(queries.includes(search)) continue
      await busquedaGoogleMaps(search)
      negocios = await getGoogleMaps()
      console.log(`Extrayendo ${search} (${negocios.length})`)
      await getData(search)
      await output.xlsx.writeFile(backupFile)
      console.log(`Extraxión exitosa de ${search})`)
    }
  }else{
    negocios = extract
    await getData()
  }
  await browser.close();

  async function busquedaGoogleMaps(search){
      await page.goto(`https://www.google.com/maps/search/${search}`);
      await page.waitForNavigation()
      let i = 0;
      while(!await page.waitForSelector('.HlvSq', {timeout: 500}).catch(() => false) && i < 30){
        await page.evaluate(() => {
          document.querySelectorAll('.Nv2PK')[document.querySelectorAll('.Nv2PK').length-1]?.scrollIntoView()
        })
        i++
      }
  }

  async function getGoogleMaps(){
    return await page.evaluate(() => {
      const links = []
      document.querySelectorAll('.Nv2PK').forEach((negocio) => links.push(negocio.children[0].href))
      if(links.length === 0) links.push(document.URL)
      return links
    })
  }

  async function getData(query){
    for(let negocio of negocios){
      if(collected.includes(negocio)) continue
      try{
        await page.goto(negocio, {timeout: waitTimeout})
        let data = await page.evaluate(() => {
          const name = document.querySelector('.fontHeadlineLarge')?.textContent?.trim()
          let web = document.querySelector('[aria-label*="Sitio web"]')?.textContent?.trim()
          let phone = document.querySelector('[aria-label*="Teléfono"]')?.textContent?.trim()
          if(web && web?.includes(' ')){web = undefined}
          if(web && !web?.includes('https://')){web = `https://${web}`}
          if(phone && phone.includes('Agregar')){phone = undefined}
          return {name, web, phone}
        })
        data = {query, link : negocio, ...data}
        if(data.web){
          await page.goto(data.web, {timeout: waitTimeout})
          await page.waitForNavigation().catch(() => {})
          const redes = await page.evaluate(() => {
            let email = document.querySelector('[href*="mailto"]')?.href.replace('mailto:', '')
            //if(!email || email === '') {email = document.querySelector('[href*="mailto"]')?.textContent}
            const facebook = document.querySelector('[href*="facebook.com"]')?.href
            const instagram = document.querySelector('[href*="instagram.com"]')?.href
            const linkedin = document.querySelector('[href*="linkedin.com"]')?.href
            return {email, facebook, instagram, linkedin}
          })
          data = {...data, ...redes}
        }
        if(!data.email && data.facebook){
          await page.goto(data.facebook, {timeout: waitTimeout})
          await page.waitForNavigation().catch(() => {})
          data.email = await page.evaluate(() => {
            const email = Array.from(document.querySelectorAll('.xieb3on ul .xu06os2'))?.filter((e) => e.textContent.includes('@'))[0]?.textContent || document.querySelector('[href*="mailto"]')?.href.replace('mailto:', '')
            return email
          })
        }
        console.log(data)
        console.log(++extraidos)
        worksheet.addRow(data)
        await output.xlsx.writeFile(outputFile)
      }catch{console.log('TimedOut')}
    }
  }
})();
