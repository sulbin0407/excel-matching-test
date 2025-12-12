// 데이터 서비스 레이어 - 나중에 MSSQL로 쉽게 변경 가능하도록 분리
import xlsx from "xlsx";
import fs from "fs";

// 간단한 워크북/시트 캐시로 반복 파일 읽기 비용 절감
// 파일 mtime이 변하면 자동 무효화
const workbookCache = new Map(); // filePath -> { workbook, mtimeMs }
const sheetCache = new Map(); // cacheKey(filePath+mtime+sheet) -> { rawData, headerRow, headers }
const SHEET_CACHE_LIMIT = 10;

function getCachedWorkbook(filePath) {
  const stat = fs.statSync(filePath);
  const cached = workbookCache.get(filePath);
  if (cached && cached.mtimeMs === stat.mtimeMs) {
    return { workbook: cached.workbook, mtimeMs: cached.mtimeMs };
  }
  const workbook = xlsx.readFile(filePath);
  workbookCache.set(filePath, { workbook, mtimeMs: stat.mtimeMs });
  return { workbook, mtimeMs: stat.mtimeMs };
}

function getCachedSheetData(filePath, sheetName) {
  const { workbook, mtimeMs } = getCachedWorkbook(filePath);
  const cacheKey = `${filePath}::${mtimeMs}::${sheetName}`;
  const cached = sheetCache.get(cacheKey);
  if (cached) {
    return { ...cached };
  }

  const sheet = workbook.Sheets[sheetName];
  if (!sheet) {
    throw new Error(`시트 "${sheetName}"를 찾을 수 없습니다.`);
  }

  const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  let headerRow = 0;

  for (let i = 0; i < Math.min(10, rawData.length); i++) {
    if (!rawData[i]) continue;
    const row = rawData[i];
    if (row[0] === "거래처명" || row[0]?.toString().includes("거래처명")) {
      headerRow = i;
      break;
    }
    if (row[0] === "비고" || row[0]?.toString().includes("비고")) {
      headerRow = i;
      break;
    }
    if (row[3] === "거래처명" || row[3]?.toString().includes("거래처명")) {
      headerRow = i;
      break;
    }
    const headerKeywords = ["전표번호", "거래처명", "통화", "잔액", "반제할금액", "만기일", "계정명", "비고", "미결발생일"];
    const keywordCount = headerKeywords.filter(keyword =>
      row.some(cell => cell && String(cell).includes(keyword))
    ).length;
    if (keywordCount >= 3) {
      headerRow = i;
      break;
    }
  }

  const headers = (rawData[headerRow] || []).map(header =>
    header !== undefined && header !== null ? String(header).trim() : ""
  );

  // 캐시 크기 제한 (간단한 FIFO)
  sheetCache.set(cacheKey, { rawData, headerRow, headers });
  if (sheetCache.size > SHEET_CACHE_LIMIT) {
    const firstKey = sheetCache.keys().next().value;
    sheetCache.delete(firstKey);
  }

  return { rawData, headerRow, headers };
}

/**
 * 엑셀 파일에서 데이터를 읽어오는 함수
 * 나중에 이 함수만 MSSQL 쿼리로 변경하면 됩니다
 * @param {string} filePath - 엑셀 파일 경로
 * @param {string} sheetName - 시트 이름
 * @param {string} userName - 필터링할 사용자 이름 (옵션)
 * @param {number} page - 페이지 번호 (1부터 시작, 옵션)
 * @param {number} limit - 페이지당 행 수 (옵션)
 */
export async function getExcelData(filePath, sheetName, userName = null, page = null, limit = null) {
  try {
    const { rawData, headerRow, headers } = getCachedSheetData(filePath, sheetName);
    const dataRows = rawData.slice(headerRow + 1);
    
    console.log(`📋 헤더 행 찾기 완료: ${headerRow}번째 행`);
    console.log(`📋 헤더 목록:`, headers.slice(0, 15).map((h, i) => `${String.fromCharCode(65 + i)}열: ${h || '(빈 헤더)'}`).join(', '));
    console.log(`📋 K열(인덱스 10) 헤더: "${headers[10] || '(빈 헤더)'}"`);
    // 계정명 관련 헤더 찾기
    const 계정명헤더인덱스들 = [];
    headers.forEach((h, idx) => {
      if (h && String(h).includes("계정명")) {
        계정명헤더인덱스들.push({ 인덱스: idx, 헤더: h, 열: String.fromCharCode(65 + idx) });
      }
    });
    console.log(`📋 계정명 관련 헤더들:`, 계정명헤더인덱스들.length > 0 ? 계정명헤더인덱스들.map(h => `${h.열}열(인덱스${h.인덱스}): "${h.헤더}"`).join(', ') : '없음');

    // 헤더를 키로 사용하여 객체 배열로 변환
    // 모든 헤더를 포함하되, 빈 헤더는 인덱스 기반으로 처리
    // 🔥 K열(인덱스 10) 계정명은 항상 Column10 키로도 저장 (E열과 K열 모두 "계정명" 헤더 충돌 방지)
    let data = dataRows.map((row, rowIndex) => {
      const obj = {};
      headers.forEach((header, idx) => {
        // 헤더가 있으면 헤더명을 키로 사용, 없으면 인덱스 기반 키 사용
        const cellValue = row[idx];
        // 빈 문자열도 유효한 값으로 처리 (빈 문자열과 undefined/null 구분)
        if (header) {
          // 헤더가 있으면 헤더명을 키로 사용
          // cellValue가 undefined나 null이 아니면 그대로 사용 (빈 문자열 포함)
          if (cellValue !== undefined && cellValue !== null) {
            obj[header] = cellValue;
          } else {
            obj[header] = "";
          }
        } else {
          // 빈 헤더는 인덱스 기반 키 사용 (예: "Column13")
          if (cellValue !== undefined && cellValue !== null) {
            obj[`Column${idx}`] = cellValue;
          } else {
            obj[`Column${idx}`] = "";
          }
        }
      });
      
      // 🔥 K열(인덱스 10) 값을 항상 Column10 키로 명시적으로 저장
      // E열과 K열 모두 "계정명" 헤더가 있어서 row["계정명"]이 E열 값일 수 있으므로
      // K열 값은 인덱스 기반 키(Column10)로 항상 접근 가능하도록 보장
      if (row.length > 10) {
        const K열값 = row[10];
        if (K열값 !== undefined && K열값 !== null) {
          obj["Column10"] = K열값;
        } else {
          obj["Column10"] = "";
        }
      } else {
        obj["Column10"] = "";
      }
      
      // 디버깅: 첫 5개 행의 K열(인덱스 10) 값 확인
      if (rowIndex < 5 && headers.length > 10) {
        const K열헤더 = headers[10] || `Column10`;
        const K열값 = row[10];
        console.log(`   [dataService] 행 ${rowIndex + 1}, K열(인덱스 10) 헤더: "${K열헤더}", 값: "${K열값 || ''}"`);
        console.log(`      obj["계정명"]: "${obj["계정명"] || ''}"`);
        console.log(`      obj["Column10"]: "${obj["Column10"] || ''}" (K열 값 보장)`);
      }
      
      return obj;
    });

    // 이름 정규화 함수 (공백/괄호 제거 등) - server.js 의 normalizeName 과 동일한 방식
    function normalizeName(value) {
      return String(value || "")
        .replace(/\s+/g, "")
        .replace(/[()]/g, "")
        .trim();
    }

    // 사용자 이름으로 필터링 (있는 경우)
    // 1차: 거래처명 계열 컬럼만 사용 (요구사항: 사용자 = 거래처명 = username)
    // 2차: 1차 결과가 0건이면, 모든 텍스트 컬럼을 대상으로 재검색 (fallback)
    if (userName) {
      const target = normalizeName(userName);

      // 사용할 수 있는 후보 컬럼들 정의
      const candidateColumnsPriority = [
        // 1순위: 거래처명 계열만 사용
        ['거래처명', '거래처', '거래처 이름']
      ];

      // 실제 파일에서 존재하는 컬럼만 추출 (대소문자 구분 없이, 부분 일치 포함)
      const availableColumns = [];
      for (const group of candidateColumnsPriority) {
        for (const colName of group) {
          // 정확히 일치하는 경우
          if (headers.includes(colName)) {
            availableColumns.push(colName);
          } else {
            // 부분 일치하는 경우 (대소문자 구분 없이)
            const foundHeader = headers.find(h => h && String(h).toLowerCase().includes(colName.toLowerCase()));
            if (foundHeader && !availableColumns.includes(foundHeader)) {
              availableColumns.push(foundHeader);
            }
          }
        }
      }
      
      // 거래처명이 포함된 모든 헤더 찾기 (fallback)
      if (availableColumns.length === 0) {
        const 거래처명헤더 = headers.filter(h => h && String(h).includes('거래처명'));
        if (거래처명헤더.length > 0) {
          availableColumns.push(...거래처명헤더);
        }
      }

      if (availableColumns.length > 0) {
        const beforeCount = data.length;
        // 🔥 거래처명 컬럼만 확인 (다른 컬럼은 확인하지 않음)
        let filtered = data.filter(row => {
          return availableColumns.some(col => {
            const 거래처명값 = row[col] || "";
            if (!거래처명값) return false;
            const candidate = normalizeName(거래처명값);
            return candidate === target || candidate.includes(target);
          });
        });

        console.log(`✅ 사용자 필터링 (거래처명만): "${userName}" (사용 컬럼: ${availableColumns.join(', ')})`);
        console.log(`   전체: ${beforeCount}개 행 → 필터링 후: ${filtered.length}개 행`);
        
        // 필터링 결과 샘플 확인
        if (filtered.length > 0 && filtered.length < 10) {
          const 샘플거래처명 = filtered.slice(0, 3).map(r => r[availableColumns[0]] || "").filter(v => v);
          console.log(`   필터링된 거래처명 샘플:`, 샘플거래처명);
        }

        data = filtered;
      } else {
        console.error(`❌ 거래처명 컬럼을 찾을 수 없습니다. 사용 가능한 컬럼: ${headers.join(', ')}`);
        console.error(`❌ username 필터링을 수행할 수 없습니다. 전체 데이터를 반환합니다.`);
        // 거래처명 컬럼이 없으면 필터링하지 않고 전체 데이터 반환
      }
    }

    // 페이지네이션 적용
    const totalRows = data.length;
    let paginatedData = data;
    
    if (page !== null && limit !== null && limit > 0) {
      const startIndex = (page - 1) * limit;
      const endIndex = startIndex + limit;
      paginatedData = data.slice(startIndex, endIndex);
    }

    return {
      headers: headers, // 빈 헤더도 포함하여 인덱스 유지
      data: paginatedData,
      totalRows: totalRows,
      page: page || 1,
      limit: limit || totalRows,
      totalPages: limit && limit > 0 ? Math.ceil(totalRows / limit) : 1
    };
  } catch (error) {
    throw new Error(`데이터 읽기 오류: ${error.message}`);
  }
}

/**
 * 사용 가능한 시트 목록 가져오기
 */
export async function getSheetNames(filePath) {
  try {
    const { workbook } = getCachedWorkbook(filePath);
    return workbook.SheetNames;
  } catch (error) {
    throw new Error(`시트 목록 읽기 오류: ${error.message}`);
  }
}







