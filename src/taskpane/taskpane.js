/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Connect database lấy dữ liệu

// // create the connection to database

const dataUser = [
  {
    "id": 11,
    "name": "Đỗ Công Kiền",
    "age": 18,
    "address": "trung quoc",
    "quantity": 1,
    "unit": "tấn",
    "group": "1"
  },
  {
    "id": 12,
    "name": "Nguyễn Thanh Lâm",
    "age": 24,
    "address": "hà nội",
    "quantity": 2,
    "unit": "tấn",
    "group": "1"
  },
  {
    "id": 13,
    "name": "Nguyễn Xuân Công",
    "age": 18,
    "address": "nam định",
    "quantity": 3,
    "unit": "tấn",
    "group": "2"
  }
]




Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
      console.log('Sorry. The tutorial add-in uses Excel.js APIs that are not available in your version of Office.');
    }


    // Assign event handlers and other initialization logic.
    document.getElementById("create-report").onclick = createReport;

    //clear table
    document.getElementById("clear-table").onclick = clearTable;

    //connect database
    document.getElementById('connect-database').onclick = connectDB;
  }
});

async function connectDB() {
  await Excel.run(async (context) => {

    //connect db
    // const connection = mysql.createConnection({
    //   host: 'localhost',
    //   user: 'root',
    //   database: 'da3'
    // });

    // (await connection).execute(
    //   'SELECT * FROM `user` where `id` = ? ', [15],
    //   (error, results, fields) => {
    //     console.log(results)
    //   }
    // )
    let results = "kiền"
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();
    let test = currentWorksheet.getRange("C1")
    test.values = results

    await context.sync();
  })
    .catch((error) => {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}

async function createReport() {
  await Excel.run(async (context) => {
    //lấy worksheet đang làm việc hiện tại
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    //xóa đường grid lines
    // currentWorksheet.showGridlines = false

    //tên công ty
    let nameCompanyRange = currentWorksheet.getRange("A1:B1")
    //quốc hiệu
    let nationalNameRange = currentWorksheet.getRange("E1:H1")
    //tiêu ngữ
    let crestRange = currentWorksheet.getRange("E2:H2")
    //tên báo cáo
    let reportNameRange = currentWorksheet.getRange("B5:G6")
    //date
    let dateRange = currentWorksheet.getRange("A2:C2")

    let date = new Date()
    let dateValue = `Ngày ${date.getDate()} Tháng ${date.getMonth()} Năm ${date.getFullYear()}`



    nationalNameRange.merge()
    crestRange.merge()
    nameCompanyRange.merge()
    reportNameRange.merge()
    dateRange.merge()

    let nameCompany = currentWorksheet.getRange("A1")
    let nationalName = currentWorksheet.getRange("E1")
    let crest = currentWorksheet.getRange("E2")
    let reportName = currentWorksheet.getRange("B5")


    currentWorksheet.getRange("A2").values = [[`${dateValue}`]]


    nameCompany.values = [["Công ty Than VN"]]
    nationalName.values = [['Cộng hòa xã hội chủ nghĩa Việt Nam']]
    crest.values = [['Độc lập - Tự do - Hạnh phúc']]
    reportName.values = [['BÁO CÁO SẢN XUẤT HÀNG NGÀY']]


    nationalName.format.horizontalAlignment = "Center"
    crest.format.horizontalAlignment = "Center"
    nameCompany.format.horizontalAlignment = "Left"
    reportName.format.horizontalAlignment = "Center"
    reportName.format.font.bold = true
    reportName.format.font.size = 18


    // nameCompany.format.autofitColumns()
    // nameCompany.format.autofitRows()






    //tạo bảng mới có header
    const expensesTable = currentWorksheet.tables.add("B10:G10", true /*hasHeaders*/);
    expensesTable.name = "Report";
    // TODO2: Queue commands to populate the table with data.
    expensesTable.getHeaderRowRange().values =
      [["STT", "Họ và tên", "Tổ sản xuất", "Sản lượng", "Đơn vị", "Ghi chú"]];

    for (let user of dataUser) {
      expensesTable.rows.add(null /*add at the end*/, [
        [`${user.id}`, `${user.name}`, `Tổ ${user.group}`, `${user.quantity}`, `${user.unit}`, ''],
      ]);
    }
    // // TODO3: Queue commands to format the table.
    // expensesTable.columns.getItemAt(3).getRange().numberFormat = [['\u20AC#,##0.00']];
    // expensesTable.getRange().format.autofitColumns();
    // expensesTable.getRange().format.autofitRows();
    // const [user, field] = await pool.execute('select * from user')

    // console.log(user)
    if (Office.context.requirements.isSetSupported("ExcelApi", "1.2")) {
      currentWorksheet.getUsedRange().format.autofitColumns();
      currentWorksheet.getUsedRange().format.autofitRows();
      currentWorksheet.getUsedRange().format.horizontalAlignment = "Center"
    }

    currentWorksheet.getRange("A2").format.horizontalAlignment = "Left"
    currentWorksheet.getRange("A2").format.font.italic = true
    await context.sync();
  })
    .catch((error) => {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}


async function clearTable() {
  await Excel.run(async (context) => {

    // TODO1: Queue table creation logic here.

    //lấy worksheet đang làm việc hiện tại
    const currentWorksheet = context.workbook.worksheets.getActiveWorksheet();

    currentWorksheet.getRange().clear()
    await context.sync();
  })
    .catch((error) => {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
}