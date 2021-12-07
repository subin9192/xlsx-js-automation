const Excel = require("exceljs");
const workbook = new Excel.Workbook();

void (async () => {
  const excel = await workbook.xlsx.readFile("before.xlsx");
  const worksheets = excel.worksheets;
  for (const worksheet of worksheets) {
    // No. 숫자에 1씩 추가
    // 2번째 줄의 2번째 칸 지정
    const noRow = worksheet.getRow(2);
    const noCell = noRow.getCell(2);
    const [id, no] = noCell.value.split("-");
    const newno = `${id}-${Number(no) + 1}`;
    noCell.value = newno;
    noRow.commit();

    // Serial Number 끝자락에 1씩 추가
    // 19번째 줄의 3번째 칸 지정
    const serialNumberRow = worksheet.getRow(19);
    const serialNumberCell = serialNumberRow.getCell(3);
    const serialNumberId = serialNumberCell.value.slice(0, 6);
    const serialNumberNo = serialNumberCell.value.slice(
      6,
      serialNumberCell.value.length
    );
    serialNumberCell.value = `${serialNumberId}${Number(serialNumberNo) + 1}`;
    serialNumberRow.commit();
    console.log(`시트명 "${worksheet.name}" 의 정보 변경 완료`);
  }

  workbook.xlsx.writeFile("after.xlsx");
  console.log("after.xlsx 파일 생성 완료");
})();
