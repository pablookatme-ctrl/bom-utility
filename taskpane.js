// Office 라이브러리 준비 이벤트 등록
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("btn-group").onclick = runBOMGrouping;
        document.getElementById("btn-reset").onclick = resetBOM;
    }
});

/**
 * 레벨에 따른 RGB 색상 반환
 */
function getColorForLevel(level) {
    switch (level) {
        case 1: return "#FFFF00"; // 노랑
        case 2: return "#FDE9D9"; // 연주황
        case 3: return "#DAEEF3"; // 연파랑
        case 4: return "#E2EFDA"; // 연초록
        case 5: return "#E6E0EC"; // 연보라
        case 6: return "#FFDCE1"; // 연분홍
        case 7: return "#CCFFCC"; // 밝은초록
        default: return "#EBEBEB"; // 회색
    }
}

/**
 * 진행률 UI 업데이트
 */
function updateProgress(show, text = "", percent = 0) {
    const container = document.getElementById("status-container");
    const textField = document.getElementById("status-text");
    const percentField = document.getElementById("status-percent");
    const fillBar = document.getElementById("progress-fill");

    if (show) {
        container.classList.remove("hidden");
        textField.innerText = text;
        percentField.innerText = `${percent}%`;
        fillBar.style.width = `${percent}%`;
    } else {
        container.classList.add("hidden");
    }
}

/**
 * [기능 1] BOM 그룹화 및 색상 적용
 */
async function runBOMGrouping() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            
            // 데이터 범위 파악을 위해 UsedRange 로드
            const usedRange = sheet.getUsedRange();
            usedRange.load(["rowCount", "address"]);
            await context.sync();

            const lastRow = usedRange.rowCount;
            if (lastRow < 3) {
                alert("처리할 데이터가 부족합니다. (최소 3행 이상 필요)");
                return;
            }

            updateProgress(true, "데이터 분석을 준비 중입니다...", 0);

            // 대량 데이터 로드 성능 최적화 (1~15열 값만 일괄 로드)
            const dataRange = sheet.getRangeByIndexes(2, 0, lastRow - 2, 15);
            dataRange.load("values");
            await context.sync();

            const values = dataRange.values;
            let baseColumn = -1;
            let foundBase = false;

            // 기존 배경 색상 리셋 (1~25열 범위)
            const targetColorRange = sheet.getRangeByIndexes(2, 0, lastRow - 2, 25);
            targetColorRange.format.fill.clear();

            // 기존 아웃라인 그룹 리셋
            try {
                const entireRows = targetColorRange.getEntireRow();
                entireRows.ungroup(Excel.GroupOption.byRows);
            } catch (e) {
                // 해제할 그룹이 없더라도 다음 로직을 위해 무시 처리
            }

            // 분석 및 변경 작업 진행
            for (let i = 0; i < values.length; i++) {
                const r = i + 2; 
                let starColumn = -1;

                for (let c = 0; c < 15; c++) {
                    if (String(values[i][c]).trim() === "*") {
                        starColumn = c;
                        break;
                    }
                }

                if (starColumn !== -1) {
                    if (!foundBase) {
                        baseColumn = starColumn;
                        foundBase = true;
                    }

                    let level = starColumn - baseColumn + 1;
                    if (level < 1) level = 1;
                    if (level > 8) level = 8;

                    const rowRange = sheet.getRangeByIndexes(r, 0, 1, 25);
                    const color = getColorForLevel(level);
                    rowRange.format.fill.color = color;

                    // 그룹화 적용
                    const entireRow = rowRange.getEntireRow();
                    for (let g = 0; g < level; g++) {
                        entireRow.group(Excel.GroupOption.byRows);
                    }
                }

                // 20행 단위로 진행률 업데이트
                if (i % 20 === 0 || i === values.length - 1) {
                    const percent = Math.round(((i + 1) / values.length) * 100);
                    updateProgress(true, `처리 중... (${i + 1} / ${values.length})`, percent);
                    await context.sync(); 
                }
            }

            await context.sync();
            updateProgress(true, "작업이 성공적으로 완료되었습니다!", 100);
            setTimeout(() => updateProgress(false), 3000); 
        });
    } catch (error) {
        console.error(error);
        alert("오류가 발생했습니다: " + error.message);
        updateProgress(false);
    }
}

/**
 * [기능 2] 초기화 함수
 */
async function resetBOM() {
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const usedRange = sheet.getUsedRange();
            usedRange.load(["rowCount"]);
            await context.sync();

            const lastRow = usedRange.rowCount;
            if (lastRow < 3) return;

            updateProgress(true, "데이터를 초기화하는 중입니다...", 50);

            const targetRange = sheet.getRangeByIndexes(2, 0, lastRow - 2, 25);
            targetRange.format.fill.clear();

            try {
                const entireRows = targetRange.getEntireRow();
                entireRows.ungroup(Excel.GroupOption.byRows);
            } catch (e) {
                // 무시
            }

            await context.sync();
            updateProgress(true, "초기화가 완료되었습니다.", 100);
            setTimeout(() => updateProgress(false), 2000);
        });
    } catch (error) {
        console.error(error);
        alert("초기화 과정 중 오류가 발생했습니다: " + error.message);
        updateProgress(false);
    }
}
