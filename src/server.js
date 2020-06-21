/* use strict */

const express = require('express')
const nodeApp = express()

const cors = require('cors')
const bodyParser = require('body-parser')

const XLSX = require('xlsx')
const XLGEN = require('xlgen')

nodeApp.use(cors())

// request entity scale
nodeApp.use(bodyParser.json({limit:'50mb'}))
nodeApp.use(bodyParser.urlencoded({limit:'50mb', extended:true}))

// function list
function alphaToNum(alpha) {
    var i = 0,
    num = 0,
    len = alpha.length;

    for (; i < len; i++) {
      num = num * 26 + alpha.charCodeAt(i) - 0x40;
    }

    return num - 1;
  }

nodeApp.post('/read', (req, res) => {
  let path = req.body.path

  let wb = XLSX.readFile(path, {raw : true})
  let ws = wb.Sheets[wb.SheetNames[0]]
  let fileType = (path).match(/\.[a-z]*$/i)[0]

  // range setting
  let range = Object.keys(ws)
  range.splice(range.indexOf('!ref'), 1)
  range.splice(range.indexOf('!margins'), 1)

  let getMaxAlpha = range.reduce( (pre, cur) => {
  return alphaToNum(pre.replace(/[^A-Z]*/g, "")) > alphaToNum(cur.replace(/[^A-Z]*/g, "")) ? pre : cur
  })

  let getMaxNumber = range.reduce( (pre, cur) => {
  return Number(pre.replace(/[^0-9]*/g, "")) > Number(cur.replace(/[^0-9]*/g, "")) ? pre : cur
  })

  let maxAlpha = getMaxAlpha.replace(/[^A-Z]*/g, "")
  let maxNumber = getMaxNumber.replace(/[^0-9]*/g, "")

  ws['!ref'] = 'A1:' + maxAlpha + maxNumber

  res.json({
    maxAlpha : maxAlpha,
    maxNumber : maxNumber,
    fileType : fileType,
    ws : ws,
    wb : wb
  })
})

nodeApp.post('/writeXLS', (req, res) => {
  let path = req.body.path
  let ws = req.body.ws

  let xlg = XLGEN.createXLGen(path)
  let sheet = xlg.addSheet('sheet')


  for(let x = 0; x < Object.keys(ws).length; x++) {
    if(alphaToNum(Object.keys(ws)[x]) > 0) {
      
      sheet.cell((Number(Object.keys(ws)[x].replace(/[^0-9]*/g, "")) - 1),
                 alphaToNum(Object.keys(ws)[x].replace(/[^A-Z]*/g, "")),
                 ws[Object.keys(ws)[x]].v)
    }
  }

  // 최종 파일 생성
  xlg.end(function(err){
    if(err) console.log(err.name, err.message);
    else console.log('writing function complete');
  });    
  
  res.end('write .xls complete')
})

nodeApp.post('/writeXLSX', (req, res) => {
    let wb = req.body.wb
    let ws = req.body.ws
    let path = req.body.path
    
    wb.Sheets[wb.SheetNames[0]] = ws
    XLSX.writeFile(wb, path, {bookType:'xlsx', type:'binary'})
    res.end('write .xlsx complete')
})

nodeApp.listen(8082, () => {
    console.log('')
})