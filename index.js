var express = require('express')

var isJson = require('is-json')

var json2xls = require('json2xls')

var bodyParser = require('body-parser')

var fs = require('fs')

const app = express()

app.use(bodyParser.json())
app.use(bodyParser.urlencoded({extended:false}))

app.set('view engine','ejs')

const PORT = 4600

app.get('/',(req,res) => {
    res.render('index',{title:'Convert Raw JSON to Excel File Free Converter - FreeMediaTools.com'})
})

app.post('/jsontoexcel',(req,res) => {

    var jsondata = req.body.json

    var exceloutput = Date.now() + "output.xlsx"

    if(isJson(jsondata)){
        var xls = json2xls(JSON.parse(jsondata));

        fs.writeFileSync(exceloutput, xls, 'binary');

        res.download(exceloutput,(err) => {
            if(err){
                fs.unlinkSync(exceloutput)
                res.send("Unable to download the excel file")
            }
            fs.unlinkSync(exceloutput)
        })
    }
    else{
        res.send("JSON Data is not valid")
    }

})

app.listen(PORT,() => {
    console.log("Your app is listening on port 4600")
})