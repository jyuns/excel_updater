/* use strict */

const express = require('express')
const nodeApp = express()

const cors = require('cors')
const bodyParser = require('body-parser')

const XLSX = require('xlsx')
const XLGEN = require('xlgen')

nodeApp.use(cors())

const getDurationInMilliseconds = (start) => {
  const NS_PER_SEC = 1e9
  const NS_TO_MS = 1e6
  const diff = process.hrtime(start)

  return (diff[0] * NS_PER_SEC + diff[1]) / NS_TO_MS
}

nodeApp.use((req, res, next) => {
  console.log(`${req.method} ${req.originalUrl} [STARTED]`)
  const start = process.hrtime()

  res.on('finish', () => {            
      const durationInMilliseconds = getDurationInMilliseconds (start)
      console.log(`${req.method} ${req.originalUrl} [FINISHED] ${durationInMilliseconds .toLocaleString()} ms`)
  })

  res.on('close', () => {
      const durationInMilliseconds = getDurationInMilliseconds (start)
      console.log(`${req.method} ${req.originalUrl} [CLOSED] ${durationInMilliseconds .toLocaleString()} ms`)
  })

  next()
})

// request entity scale
nodeApp.use(bodyParser.json({limit:'100mb'}))
nodeApp.use(bodyParser.urlencoded({limit:'100mb', extended:true}))

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

  let wb = XLSX.readFile(path)
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

  let key = Object.keys(ws)

  for(let x = 0; x < key.length; x++) {
    if(alphaToNum(key[x]) > 0) {
      sheet.cell((Number(key[x].replace(/[^0-9]*/g, "")) - 1),
                 alphaToNum(key[x].replace(/[^A-Z]*/g, "")),
                 ws[key[x]].v)
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
  let ws = req.body.ws
  let path = req.body.path

  let wb = XLSX.utils.book_new()
  
  XLSX.utils.book_append_sheet(wb, ws)

  XLSX.writeFile(wb, path, {bookSST:true, compression:true, ignoreEC : false, bookVBA:true, type:"binary", bookType:'xlsx'})
  res.end('write .xlsx complete')
})


nodeApp.listen(8082, () => {
    console.log('')
})