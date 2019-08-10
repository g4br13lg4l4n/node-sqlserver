const createCsvWriter = require('csv-writer').createObjectCsvWriter
const Sequelize = require('sequelize')
const XLSX = require('xlsx')


const csvWriter = createCsvWriter({
  path: 'excel.file',
  header: [
    { id: 'vin', title: 'vin' },
    { id: 'status', title: 'status'}
  ]
})


const config = {
  DB_HOST: '',
  DB_NAME: '',
  DB_USER: '',
  DB_PASSWORD: ''
}

var sequelize

sequelize = new Sequelize(config.DB_NAME, config.DB_USER, config.DB_PASSWORD, {
  host: config.DB_HOST,
  dialect: 'mssql',
  dialectOptions: {
    options: {
      encrypt: true
    },
    requestTimeout: 30000,
    trustServerCertificate: false,
    hostNameInCertificate: '*.database.windows.net'
  },
  logging: false
})


const Vehicle = sequelize.define('vehicle', {
  external_id: Sequelize.TEXT,
  external_user_id: Sequelize.TEXT,

  hologram: Sequelize.INTEGER,
  brand: Sequelize.TEXT,
  model: Sequelize.TEXT,
  name: Sequelize.TEXT,
  plate: Sequelize.TEXT,
  subbrand: Sequelize.TEXT,
   vin: Sequelize.TEXT,
})


const workbook = XLSX.readFile('VINES.xlsx')
const sheet_name_list = workbook.SheetNames
let readXlsx = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]])

console.log(readXlsx.length)


var i
for (i = 76001; i <= 77000; i++) { 
  let _vin = readXlsx[i].VIN

  sequelize.models.vehicle.findOne({
    where: { vin: _vin },
    attributes: ['vin']
  }).then(resp => {

    if(resp){
      console.log({ vin :_vin, status: 'Si' })
      csvWriter.writeRecords([{ vin :_vin, status: 'Si' }])
    }else {
      console.log({ vin :_vin, status: 'No' })
      csvWriter.writeRecords([{ vin: _vin, status: 'No'}])
    }
  })
  .catch(err => {
    console.log('err -->', err)
  })

}


/*
const vins =  async () => readXlsx.map(e => {
  

   sequelize.models.vehicle.findOne({
      where: { vin: _vin },
      attributes: ['vin']
    }).then(resp => {

      if(resp){
        console.log({ vin :_vin, status: 'Si' })
        csvWriter.writeRecords([{ vin :_vin, status: 'Si' }])
      }else {
        //console.log({ vin :_vin, status: 'No' })
        //csvWriter.writeRecords([{ vin: _vin, status: 'No'}])
      }
    })
    .catch(err => {
      console.log('err -->', err)
    })
    console.log(cont ++)
})
*/



 

