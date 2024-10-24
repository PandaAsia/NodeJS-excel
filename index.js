import XlsxPopulate from "xlsx-populate";

// XlsxPopulate.fromBlankAsync().then((Workbook) => {
//   Workbook.sheet(0).cell("A1").value("hello world");
//   return Workbook.toFileAsync("./salida.xlsx");
// });

async function createExcel() {
  const workbook = await XlsxPopulate.fromBlankAsync();
  workbook.sheet(0).cell("A1").value("nombre");
  workbook.sheet(0).cell("B1").value("apellido");
  workbook.sheet(0).cell("C1").value("Edad");

  workbook.sheet(0).cell("A2").value("shana");
  workbook.sheet(0).cell("B2").value("yamato");
  workbook.sheet(0).cell("C2").value(15);

  workbook.toFileAsync("./salida2.xlsx");
}

async function ReadExcel() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida2.xlsx");
  const value = workbook.sheet("Sheet1").cell("A2").value();
  console.log(value);
}

async function ReadAllExcel() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida2.xlsx");
  const value = workbook.sheet("Sheet1").usedRange().value();
  console.log(value);
}

async function ReadAllExcel2() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida2.xlsx");
  const value = workbook.sheet("Sheet1").range("A1:A2").value();
  console.log(value);
}

async function createRanger() {
  const workbook = await XlsxPopulate.fromBlankAsync();
  workbook
    .sheet(0)
    .cell("A1")
    .value([
      ["nombre", "sexo"],
      ["shana", "si"],
      ["jerry", "no"],
    ]);
  workbook.toFileAsync("./salida.xlsx");
}

async function editExcel() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida.xlsx");
  // const sheet = workbook.sheet(0);
  // console.log(sheet.name())

  workbook.addSheet("hoja 2");
  workbook.toFileAsync("./salida.xlsx");
}

async function editExcel2() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida.xlsx");
  workbook.sheets().map((el) => el.range);
}

async function editExcel3() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida.xlsx");
  workbook.sheet("hoja 2").name("hoja 127");
  workbook.toFileAsync("./salida.xlsx");
}

async function deleteExcel3() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida.xlsx");

  workbook.deleteSheet("hoja 2");
}

async function passwordExcel3() {
  const workbook = await XlsxPopulate.fromFileAsync("./salida.xlsx");
  workbook.addSheet("seguro");
  workbook.toFileAsync("./salida.xlsx", { password: "123" });
}
deleteExcel3();
