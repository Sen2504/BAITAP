function handleFile(){

const file=document.getElementById("excelFile").files[0]

if(!file){
alert("Chọn file Excel trước")
return
}

const reader=new FileReader()

reader.onload=function(e){

const data=new Uint8Array(e.target.result)

const workbook=XLSX.read(data,{type:"array"})

const sheet=workbook.Sheets[workbook.SheetNames[0]]

const json=XLSX.utils.sheet_to_json(sheet,{header:1})

renderTable(json)

createChart(json)

}

reader.readAsArrayBuffer(file)

}



function renderTable(data){

let html="<table>"

data.forEach((row,i)=>{

html+="<tr>"

row.forEach(cell=>{

if(i==0)
html+=`<th>${cell??""}</th>`
else
html+=`<td>${cell??""}</td>`

})

html+="</tr>"

})

html+="</table>"

document.getElementById("tableArea").innerHTML=html

}



function createChart(data){

let labels=[]
let values=[]

for(let i=1;i<data.length;i++){

if(!data[i][1]) continue

labels.push(data[i][1])

values.push(Number(data[i][5])||0)

}

const ctx=document.getElementById("chart")

new Chart(ctx,{
type:"pie",
data:{
labels:labels,
datasets:[{
data:values
}]
}
})

}