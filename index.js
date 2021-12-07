// Exceljs 오픈소스 초기화 과정. 건들 필요 없음.
const Excel = require("exceljs");
const workbook = new Excel.Workbook();

void (async () => {
  const excel = await workbook.xlsx.readFile("before.xlsx"); // 파일을 읽어온다.
  const worksheets = excel.worksheets; // 시트들을 가져온다.

  // 첫 번째 시트정보 얻어오기
  const originSheet = worksheets[0]; // 0번째를 가져온다. (필수)

  // 기존 시트 전체 삭제 (불필요한 시트 삭제. 없으면 적지 않아도 됨.)
  for (const worksheet of worksheets)
    await excel.removeWorksheet(worksheet.name);

  // 새로운 시트 300개 생성 (만들 시트의 개수를 기입한다.)
  for (let index = 1; index <= 300; index++) {
    const worksheet = await excel.addWorksheet(`Sheet ${index}`, {
      // 기존 시트 설정 복사
      headerFooter: originSheet.headerFooter,
      pageSetup: originSheet.pageSetup,
      properties: originSheet.properties,
      views: originSheet.views,
      state: originSheet.state,
    });

    // 기존 시트 정보 복사
    worksheet.model = Object.assign(originSheet.model, {
      mergeCells: originSheet.merges,
    });

    // 각 셀 스타일까지 모두 복사하기
    originSheet.eachRow((row, rowNumber) => {
      const newRow = worksheet.getRow(rowNumber);
      row.eachCell((cell, colNumber) => {
        const newCell = newRow.getCell(colNumber);
        for (const prop in cell) newCell[prop] = cell[prop];
      });
    });

    worksheet.name = `Sheet ${index}`;

    // No. 숫자에 1씩 추가
    // 2번째 줄의 2번째 칸 지정
    const noRow = worksheet.getRow(2);
    const noCell = noRow.getCell(2);
    const [id, no] = noCell.value.split("-");
    noCell.value = `${id}-${Number(no) + (index - 1)}`;
    noRow.commit();

    // Serial Number 끝자락에 1씩 추가
    // 19번째 줄의 3번째 칸 지정
    const serialNumberRow = worksheet.getRow(19);
    const serialNumberCell = serialNumberRow.getCell(3);
    const serialNumberId = serialNumberCell.value.slice(0, 6); // 공백 포함 총 6글자
    const serialNumberNo = serialNumberCell.value.slice(
      6,
      serialNumberCell.value.length
    );  // 숫자만 잘라온다.
    serialNumberCell.value = `${serialNumberId}${
      Number(serialNumberNo) + (index - 1)
    }`;
    serialNumberRow.commit();

    console.log(`시트명 "${worksheet.name}" 생성 완료`);
  }

  await workbook.xlsx.writeFile("after.xlsx");
  console.log("after.xlsx 파일 생성 완료");
})();
