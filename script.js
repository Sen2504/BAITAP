let myChart = null
Chart.register(ChartDataLabels)

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

let json = XLSX.utils.sheet_to_json(sheet,{header:1})

json = fixMergedCells(json)
json = removeColumn(json,3)

function removeColumn(data,colIndex){

data.forEach(row=>{
row.splice(colIndex,1)
})
console.log(data)
return data

}
renderTable(json)
createChart(json)

}

reader.readAsArrayBuffer(file)
}

function fixMergedCells(data){

for(let r=0;r<data.length;r++){

for(let c=1;c<data[r].length;c++){

if(data[r][c]===undefined){
data[r][c]=data[r][c-1]
}

}

}

return data
}

function renderTable(data){

let html="<table>"

/* bỏ phần header Excel */
let tableData = data.slice(6, data.length-5)

tableData.forEach((row,i)=>{

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

/* bỏ header excel */
let rows = data.slice(6, data.length-5)

/* lấy tên khoa + số người */
let labels=[]
let values=[]

    rows.forEach(row=>{

    let khoa=row[2]
    let soNguoi=parseFloat(row[5])

    if(khoa && !isNaN(soNguoi) && soNguoi > 0 && khoa.includes("Khoa") ){

    labels.push(khoa)
    values.push(soNguoi)

    }

    })

    // console.log(labels)
    // console.log(values)
    // console.log(rows[6])
    // console.log(rows[7])

/* màu tự động */
let colors=[
"#ff6384",
"#36a2eb",
"#ffce56",
"#4bc0c0",
"#9966ff",
"#ff9f40",
"#2ecc71",
"#e74c3c"
]

/* nếu khoa nhiều hơn màu → random thêm */
while(colors.length < labels.length){

colors.push(
'#'+Math.floor(Math.random()*16777215).toString(16)
)

}

const ctx=document.getElementById("chart").getContext("2d")

/* xoá chart cũ */
if(myChart){
myChart.destroy()
}

myChart = new Chart(ctx,{
type:"pie",
data:{
labels:labels,
datasets:[{
data:values,
backgroundColor:colors
}]
},
options:{
responsive:true,
plugins:{
legend:{
display:false
},
datalabels:{
color:"#fff",
font:{
weight:"bold",
size:14
},
formatter:(value,context)=>{

let total=context.chart.data.datasets[0].data.reduce((a,b)=>a+b,0)
let percent=((value/total)*100).toFixed(1)

return value+" người\n"+percent+"%"

}
}
}
},
plugins:[ChartDataLabels]

})

createLegend(labels,colors)

}

function createLegend(labels,colors){

let html=""

labels.forEach((label,i)=>{

html+=`
<div class="legend-item">
<div class="legend-color" style="background:${colors[i]}"></div>
<span>${label}</span>
</div>
`

})

document.getElementById("chartLegend").innerHTML=html

}