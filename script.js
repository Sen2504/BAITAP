let myChart = null
let myChart2 = null
let myChart3 = null
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
json = removeColumn(json,12)

function removeColumn(data,colIndex){

data.forEach(row=>{
row.splice(colIndex,1)
})
console.log(data)
return data

}

renderTable(json)
createChart(json)
createDeathChart(json)
createBarChart(json)
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

function createChart(data) {

  let rows = data.slice(8, data.length - 5)

  let khoaData = []

  rows.forEach(row => {

    let khoa = row[2]
    let treEm = parseFloat(row[6]) || 0
    let capCuu = parseFloat(row[7]) || 0

    if (
      khoa &&
      typeof khoa === "string" &&
      khoa.includes("Khoa") &&
      khoa !== "Tổng số" &&
      khoa !== "Nữ"
    ) {
      khoaData.push({
        khoa: khoa,
        treEm: treEm,
        capCuu: capCuu
      })
    }
  })

  console.log("khoaData thật từ Excel:", khoaData)

  const outerLabels = []
  const outerValues = []

  const innerLabels = []
  const innerValues = []

  const outerColors = [
    "#ffb703",
    "#219ebc",
    "#8e6bbd",
    "#fb8500",
    "#90be6d",
    "#f94144",
    "#577590"
  ]

  const innerColors = [
    "#ffd166", "#fcbf49",
    "#8ecae6", "#48cae4",
    "#cdb4db", "#b98ad7",
    "#ffafcc", "#a0c4ff",
    "#caffbf", "#fdffb6",
    "#f4a261", "#e76f51",
    "#84a59d", "#f28482"
  ]

  khoaData.forEach(item => {
    const tong = item.treEm + item.capCuu

    if (tong > 0) {
      outerLabels.push(item.khoa)
      outerValues.push(tong)

      if (item.treEm > 0) {
        innerLabels.push(item.khoa + " - Trẻ em < 15")
        innerValues.push(item.treEm)
      }

      if (item.capCuu > 0) {
        innerLabels.push(item.khoa + " - Số cấp cứu")
        innerValues.push(item.capCuu)
      }
    }
  })

  if (outerValues.length === 0) {
    alert("Không có dữ liệu để vẽ biểu đồ")
    return
  }

  const ctx = document.getElementById("chart").getContext("2d")

  if (myChart) {
    myChart.destroy()
  }

  myChart = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: [...outerLabels, ...innerLabels],
      datasets: [
        {
          label: "Tổng theo khoa",
          data: outerValues,
          backgroundColor: outerColors.slice(0, outerValues.length),
          borderWidth: 1,
          weight: 2
        },
        {
          label: "Chi tiết",
          data: innerValues,
          backgroundColor: innerColors.slice(0, innerValues.length),
          borderWidth: 1,
          weight: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "35%",
      plugins: {
        legend: {
          display: false
        },
        datalabels: {
          color: "#000",
          font: {
            weight: "bold",
            size: 10
          },
          formatter: (value, context) => {
            const tongTatCa = context.chart.data.datasets[1].data.reduce((a, b) => a + b, 0)
            const percent = ((value / tongTatCa) * 100).toFixed(1)
            return value + "\n" + percent + "%"
          }
        }
      }
    },
    plugins: [ChartDataLabels]
  })

  createLegend(khoaData, outerColors)
}

function createLegend(khoaData, colors) {

  let html = ""

  khoaData.forEach((item, i) => {

    const tong = item.treEm + item.capCuu

    if (tong > 0) {
      html += `
      <div class="legend-item">
        <div class="legend-color" style="background:${colors[i]}"></div>
        <span>${item.khoa} (Tổng: ${tong})</span>
      </div>
      `
    }

  })

  document.getElementById("chartLegend").innerHTML = html
}

function createDeathChart(data) {

  let rows = data.slice(8, data.length - 5)

  let khoaData = []

  rows.forEach(row => {

    let khoa = row[2]

    // ===== THAY 2 CỘT NÀY THEO FILE EXCEL CỦA MÀY =====
    let tuVongTreEm = parseFloat(row[10]) || 0
    let tuVongNguoiLon = parseFloat(row[11]) || 0
    // ================================================

    if (
      khoa &&
      typeof khoa === "string" &&
      khoa.includes("Khoa") &&
      khoa !== "Tổng số" &&
      khoa !== "Nữ"
    ) {
      khoaData.push({
        khoa: khoa,
        tuVongTreEm: tuVongTreEm,
        tuVongNguoiLon: tuVongNguoiLon
      })
    }

  })

  console.log("khoaData tử vong:", khoaData)

  const outerLabels = []
  const outerValues = []

  const innerLabels = []
  const innerValues = []

  const outerColors = [
    "#ff6b6b",
    "#4dabf7",
    "#9775fa",
    "#ffa94d",
    "#69db7c",
    "#f06595",
    "#74c0fc"
  ]

  const innerColors = [
    "#ffc9c9", "#ffa8a8",
    "#a5d8ff", "#74c0fc",
    "#d0bfff", "#b197fc",
    "#ffd8a8", "#ffbe6f",
    "#b2f2bb", "#8ce99a",
    "#fcc2d7", "#faa2c1",
    "#c5f6fa", "#99e9f2"
  ]

  khoaData.forEach(item => {
    const tong = item.tuVongTreEm + item.tuVongNguoiLon

    if (tong > 0) {
      outerLabels.push(item.khoa)
      outerValues.push(tong)

      if (item.tuVongTreEm > 0) {
        innerLabels.push(item.khoa + " - Tử vong trẻ em")
        innerValues.push(item.tuVongTreEm)
      }

      if (item.tuVongNguoiLon > 0) {
        innerLabels.push(item.khoa + " - Tử vong người lớn")
        innerValues.push(item.tuVongNguoiLon)
      }
    }
  })

  if (outerValues.length === 0) {
    document.getElementById("chartLegend2").innerHTML = "<p>Không có dữ liệu tử vong để vẽ</p>"
    return
  }

  const ctx = document.getElementById("chart2").getContext("2d")

  if (myChart2) {
    myChart2.destroy()
  }

  myChart2 = new Chart(ctx, {
    type: "doughnut",
    data: {
      labels: [...outerLabels, ...innerLabels],
      datasets: [
        {
          label: "Tổng tử vong theo khoa",
          data: outerValues,
          backgroundColor: outerColors.slice(0, outerValues.length),
          borderWidth: 1,
          weight: 2
        },
        {
          label: "Chi tiết tử vong",
          data: innerValues,
          backgroundColor: innerColors.slice(0, innerValues.length),
          borderWidth: 1,
          weight: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      cutout: "35%",
      plugins: {
        legend: {
          display: false
        },
        datalabels: {
          color: "#000",
          font: {
            weight: "bold",
            size: 10
          },
          formatter: (value, context) => {
            const dataset = context.chart.data.datasets[1].data
            const tongTatCa = dataset.reduce((a, b) => a + b, 0)
            const percent = ((value / tongTatCa) * 100).toFixed(1)
            return value + "\n" + percent + "%"
          }
        }
      }
    },
    plugins: [ChartDataLabels]
  })

  createDeathLegend(khoaData, outerColors)
}

function createDeathLegend(khoaData, colors) {

  let html = ""

  khoaData.forEach((item, i) => {

    const tong = item.tuVongTreEm + item.tuVongNguoiLon

    if (tong > 0) {
      html += `
      <div class="legend-item">
        <div class="legend-color" style="background:${colors[i]}"></div>
        <span>${item.khoa} (Tử vong: ${tong})</span>
      </div>
      `
    }

  })

  document.getElementById("chartLegend2").innerHTML = html
}

function createBarChart(data) {
  let rows = data.slice(8, data.length - 5)

  let labels = []
  let dauKyData = []
  let cuoiKyData = []

  let tongDauKy = 0
  let tongCuoiKy = 0

  rows.forEach(row => {
    let khoa = row[2]

    // sửa 2 index này theo đúng cột trong Excel
    let dauKy = parseFloat(row[4]) || 0
    let cuoiKy = parseFloat(row[13]) || 0

    if (
      khoa &&
      typeof khoa === "string" &&
      khoa.includes("Khoa") &&
      khoa !== "Tổng số" &&
      khoa !== "Nữ"
    ) {
      labels.push(khoa)
      dauKyData.push(dauKy)
      cuoiKyData.push(cuoiKy)

      tongDauKy += dauKy
      tongCuoiKy += cuoiKy
    }
  })

  console.log("labels chart 3:", labels)
  console.log("đầu kỳ:", dauKyData)
  console.log("cuối kỳ:", cuoiKyData)
  console.log("tổng đầu kỳ:", tongDauKy)
  console.log("tổng cuối kỳ:", tongCuoiKy)

  if (labels.length === 0) {
    alert("Không có dữ liệu để vẽ biểu đồ cột")
    return
  }

  // render note custom ở góc phải
  const legendBox = document.getElementById("chartLegend3")
  legendBox.innerHTML = `
    <div class="legend-item">
      <span class="legend-color legend-blue"></span>
      <span>Số bệnh nhân đầu kỳ (Tổng: ${tongDauKy})</span>
    </div>
    <div class="legend-item">
      <span class="legend-color legend-yellow"></span>
      <span>Số bệnh nhân còn lại cuối kỳ (Tổng: ${tongCuoiKy})</span>
    </div>
  `

  const ctx = document.getElementById("chart3").getContext("2d")

  if (myChart3) {
    myChart3.destroy()
  }

  myChart3 = new Chart(ctx, {
    type: "bar",
    data: {
      labels: labels,
      datasets: [
        {
          label: "Số bệnh nhân đầu kỳ",
          data: dauKyData,
          backgroundColor: "#219ebc",
          borderWidth: 1
        },
        {
          label: "Số bệnh nhân còn lại cuối kỳ",
          data: cuoiKyData,
          backgroundColor: "#ffb703",
          borderWidth: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: {
        padding: {
          top: 10,
          right: 10,
          bottom: 0,
          left: 10
        }
      },
      plugins: {
        legend: {
          display: false
        },
        datalabels: {
          anchor: "end",
          align: "top",
          color: "#000",
          font: {
            weight: "bold",
            size: 10
          },
          formatter: function(value) {
            return value
          }
        }
      },
      scales: {
        x: {
          ticks: {
            maxRotation: 45,
            minRotation: 30,
            autoSkip: false
          }
        },
        y: {
          beginAtZero: true
        }
      }
    },
    plugins: [ChartDataLabels]
  })
}