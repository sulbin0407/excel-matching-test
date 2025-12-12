// 엑셀 K열 매칭 처리 함수 (재사용 가능)
import xlsx from "xlsx";
import stringSimilarity from "string-similarity";
import dotenv from "dotenv";
// OpenAI import 제거 (K열 계정명 매칭에서 OpenAI 사용 안 함)

// 환경변수 로드
dotenv.config();

/**
 * 엑셀 파일을 처리하여 K열을 채우는 함수
 * @param {string} inputFile - 입력 파일 경로 (기본: match_data2.xlsx)
 * @param {string} outputFile - 출력 파일 경로 (기본: match_data2_result.xlsx)
 * @returns {Promise<Object>} 처리 결과 통계
 */
export async function processExcelFile(inputFile = "match_data2.xlsx", outputFile = "match_data2_result.xlsx") {
  console.log(`\n🔄 엑셀 파일 처리 시작: ${inputFile}`);
  console.log(`   출력 파일: ${outputFile}`);
  console.log(`   시간: ${new Date().toLocaleString('ko-KR')}\n`);

  try {
      // 🔥 파일 종류별 시트 이름 설정
      // - moca  : 2024moca / 2025moca
      // - 기타  : 2024     / 2025
      const lowerInput = String(inputFile || "").toLowerCase();
      const isMocaFile = lowerInput.includes("moca");

      const learningSheetName   = isMocaFile ? "2024moca" : "2024";
      const processingSheetName = isMocaFile ? "2025moca" : "2025";
      const resultSheetName     = isMocaFile ? "2025moca" : "2025"; // 결과 파일 저장 시 시트 이름
      
      // 기존 결과 파일에서 K열 값과 I열 값 읽기 (보존용)
      // I열 값은 원본 파일의 이전 I열 값과 비교하기 위해 저장
      let 기존K열값맵 = new Map(); // 인덱스 -> K열 값
      let 기존I열값맵 = new Map(); // 인덱스 -> I열 값 (원본 파일의 이전 I열 값 추정용)
      
      try {
        const 기존워크북 = xlsx.readFile(outputFile);
        const 기존시트2025 = 기존워크북.Sheets[resultSheetName] || 기존워크북.Sheets["2025"];
        if (기존시트2025) {
          const 기존데이터2025 = xlsx.utils.sheet_to_json(기존시트2025, { header: 1, defval: "" });
          // 헤더 행 찾기
          let 기존헤더행 = 0;
          for (let i = 0; i < Math.min(10, 기존데이터2025.length); i++) {
            const row = 기존데이터2025[i] || [];
            const firstCell = String(row[0] || "").trim();
            if (firstCell.includes("비고") || firstCell.includes("적요")) {
              기존헤더행 = i;
              break;
            }
          }
          
          // 기존 결과 파일의 K열 값과 I열 값 저장
          const 기존데이터행 = 기존데이터2025.slice(기존헤더행 + 1);
          const 기존헤더 = 기존데이터2025[기존헤더행] || [];
          const 기존K열인덱스 = 기존헤더.findIndex(h => {
            const hStr = String(h || "").trim();
            return hStr === "계정명" || hStr === "사용처" || hStr.includes("K");
          });
          const 기존I열인덱스 = 기존헤더.findIndex(h => String(h || "").includes("비고") || String(h || "").includes("I"));
          
          if (기존K열인덱스 !== -1 && 기존I열인덱스 !== -1) {
            기존데이터행.forEach((row, idx) => {
              const 기존K값 = String(row[기존K열인덱스] || "").trim();
              const 기존I값 = String(row[기존I열인덱스] || "").trim();
              // K열 값은 "-" 포함하여 모두 저장 (I열이 변경되지 않았으면 기존 값 유지)
              if (기존K값 !== undefined && 기존K값 !== null) {
                기존K열값맵.set(idx, 기존K값);
              }
              // I열 값도 저장 (원본 파일의 이전 I열 값과 비교용)
              // 주의: 기존 결과 파일의 I열 값은 원본 파일의 I열 값과 동일해야 함
              // 원본 파일이 변경되지 않았다면 동일함
              if (기존I값 && 기존I값 !== "") {
                기존I열값맵.set(idx, 기존I값);
              }
            });
          }
          
          if (기존K열값맵.size > 0) {
            console.log(`   📋 기존 K열 값 보존: ${기존K열값맵.size}개 행`);
          }
        }
      } catch (error) {
        // 결과 파일이 없으면 모든 행을 새로 처리
        console.log(`   📋 결과 파일 없음: 모든 행을 새로 처리합니다.`);
        // 첫 번째 실행이므로 기존 값 맵은 비어있음
        기존K열값맵 = new Map();
        기존I열값맵 = new Map();
      }

    // match_data2.xlsx 파일 읽기
    const workbook = xlsx.readFile(inputFile);
    
    // 🔥 moca 파일인 경우 2024moca/2025moca 시트 사용, 그 외에는 2024/2025 시트 사용
    // (위에서 이미 선언했으므로 재선언하지 않음)
    const sheet2024 = workbook.Sheets[learningSheetName] || workbook.Sheets["2024"];
    const sheet2025 = workbook.Sheets[processingSheetName] || workbook.Sheets["2025"];

    if (!sheet2024 || !sheet2025) {
      throw new Error(`❌ 오류: ${learningSheetName} 또는 ${processingSheetName} 시트를 찾을 수 없습니다. 사용 가능한 시트: ${workbook.SheetNames.join(', ')}`);
    }
    
    if (isMocaFile) {
      console.log(`📋 moca 파일 감지:`);
      console.log(`   - 학습 데이터: ${learningSheetName} 시트`);
      console.log(`   - 처리 대상: ${processingSheetName} 시트`);
    }

    // 시트 데이터를 배열로 읽기 (헤더 포함)
    // 🔥 원본 시트의 !ref 범위 확인 및 확장 (모든 행 포함 보장)
    const sheet2025Range = sheet2025['!ref'] ? xlsx.utils.decode_range(sheet2025['!ref']) : null;
    if (sheet2025Range && isMocaFile) {
      console.log(`   📊 원본 시트 범위: ${sheet2025['!ref']} (행 ${sheet2025Range.e.r + 1}까지)`);
      console.log(`   ⚠️ 예상 행 수: 21953행 (헤더 제외)`);
      console.log(`   ⚠️ 실제 범위 행 수: ${sheet2025Range.e.r + 1}행 (헤더 포함)`);
    }
    
    const data2024 = xlsx.utils.sheet_to_json(sheet2024, { header: 1, defval: "" });
    const data2025 = xlsx.utils.sheet_to_json(sheet2025, { header: 1, defval: "" });
    
    if (isMocaFile) {
      console.log(`   📊 읽은 데이터 행 수: ${data2025.length}개 (헤더 포함)`);
      console.log(`   ⚠️ 예상 행 수와 비교: ${data2025.length}개 vs ${(sheet2025Range?.e.r || 0) + 1}개 (범위)`);
      if (data2025.length < 21953) {
        console.log(`   ❌ 경고: 읽은 행 수가 예상보다 적습니다! 원본 파일이 열려있거나 !ref 범위가 잘못되었을 수 있습니다.`);
      }
    }

    // 헤더 행 찾기
    let headerRow2024 = 0;
    let headerRow2025 = 0;

    // 2024 시트에서 헤더 행 찾기 (적요 또는 계정명이 있는 행)
    for (let i = 0; i < Math.min(10, data2024.length); i++) {
      const row = data2024[i] || [];
      const firstCell = String(row[0] || "").trim();
      if (firstCell.includes("적요") || firstCell.includes("계정명")) {
        headerRow2024 = i;
        break;
      }
    }

    // 2025 시트에서 헤더 행 찾기
    for (let i = 0; i < Math.min(10, data2025.length); i++) {
      const row = data2025[i] || [];
      const firstCell = String(row[0] || "").trim();
      if (firstCell.includes("비고") || firstCell.includes("적요")) {
        headerRow2025 = i;
        break;
      }
    }

    const header2024 = data2024[headerRow2024] || [];
    const header2025 = data2025[headerRow2025] || [];

    // 헤더 인덱스 찾기
    // 🔥 moca 파일인 경우 명확한 열 위치 사용
    let 적요Index2024, 계정명Index2024, 비고Index2025, K열Index2025, M열Index2025;
    
    if (isMocaFile) {
      // 2024moca: A열(0) = 적요, B열(1) = 계정명
      적요Index2024 = 0;
      계정명Index2024 = 1;
      
      // 2025moca: I열(8) = 비고, K열(10) = 계정명, M열(12) = 합계잔액시산표 계정명
      비고Index2025 = 8;
      K열Index2025 = 10;
      M열Index2025 = 12;
      
      console.log("   🔥 moca 파일: 고정 열 인덱스 사용");
    } else {
      // 기존 파일: 헤더에서 동적으로 찾기
      적요Index2024 = header2024.findIndex(h => String(h || "").includes("적요"));
      계정명Index2024 = header2024.findIndex(h => String(h || "").includes("계정명"));
      비고Index2025 = header2025.findIndex(h => String(h || "").includes("비고") || String(h || "").includes("I"));
      
      // K열 인덱스 찾기 - K열은 우리가 찾은 값을 넣는 계정명 열
      const 사용처Index = header2025.findIndex(h => String(h || "").trim() === "사용처");
      K열Index2025 = 10; // 기본값 (K열, 인덱스 10)

      if (사용처Index !== -1) {
        K열Index2025 = 사용처Index + 1;
      } else {
        const 계정명인덱스들 = [];
        header2025.forEach((h, idx) => {
          if (String(h || "").trim() === "계정명") {
            계정명인덱스들.push(idx);
          }
        });
        if (계정명인덱스들.length >= 2) {
          K열Index2025 = 계정명인덱스들[1];
        } else {
          K열Index2025 = 10;
        }
      }

      M열Index2025 = header2025.findIndex(h => String(h || "").includes("합계잔액시산표 계정명"));
      if (M열Index2025 === -1) {
        M열Index2025 = 12; // 기본값
      }
    }

    console.log("🔍 헤더 정보:");
    console.log(`   2024 시트 - 적요 인덱스: ${적요Index2024}, 계정명 인덱스: ${계정명Index2024}`);
    console.log(`   2025 시트 - 비고 인덱스: ${비고Index2025}, K열 인덱스: ${K열Index2025}, M열 인덱스: ${M열Index2025}`);

    // 2024 시트 데이터 파싱 (적요와 계정명의 관계 학습)
    const dataRows2024 = data2024.slice(headerRow2024 + 1);
    const 학습데이터 = [];

    dataRows2024.forEach((row, index) => {
      const 적요 = String(row[적요Index2024] || "").trim();
      const 계정명 = String(row[계정명Index2024] || "").trim();
      
      if (적요 && 계정명) {
        학습데이터.push({
          적요: 적요,
          계정명: 계정명,
          원본인덱스: index
        });
      }
    });

    console.log(`\n📚 학습 데이터: ${학습데이터.length}개 행`);

    // 날짜 형식 제거 함수
    function removeDates(text) {
      if (!text) return "";
      return String(text)
        .replace(/\d{2,4}년\s*\d{1,2}월/g, "")
        .replace(/\d{2,4}\.\d{1,2}/g, "")
        .replace(/\d{4}-\d{2}-\d{2}/g, "")
        .replace(/\d{8}/g, "")
        .replace(/\d{4}년/g, "")
        .replace(/\d{1,2}월/g, "");
    }

    // 텍스트 정규화 함수 (날짜 제거, 띄어쓰기, 기호 제거)
    function normalizeText(text) {
      if (!text) return "";
      let normalized = String(text);
      normalized = removeDates(normalized);
      normalized = normalized.replace(/\s+/g, "");
      normalized = normalized.replace(/[^\w가-힣]/g, "");
      normalized = normalized.toLowerCase();
      return normalized;
    }

    // 2025 시트 데이터 처리
    const dataRows2025 = data2025.slice(headerRow2025 + 1);
    
    if (isMocaFile) {
      console.log(`   📊 처리할 데이터 행 수: ${dataRows2025.length}개 (헤더 제외)`);
      console.log(`   📊 헤더 행 인덱스: ${headerRow2025} (${headerRow2025 + 1}번째 행)`);
    }
    
    let processedCount = 0;
    let matchedCount = 0;
    let noMatchCount = 0;
    // ai호출횟수 제거 (OpenAI 사용 안 함)

    // 2024 시트의 적요 목록 생성 (매칭용)
    const 적요목록2024 = 학습데이터.map(d => d.적요);
    const 정규화된적요목록2024 = 적요목록2024.map(적요 => normalizeText(적요));

    // 2025 시트의 M열 전체 데이터 수집 (중복 제거)
    const M열전체데이터 = [];
    dataRows2025.forEach(row => {
      const m값 = String(row[M열Index2025] || "").trim();
      if (m값 && m값 !== "" && m값 !== "-" && !M열전체데이터.includes(m값)) {
        M열전체데이터.push(m값);
      }
    });

    console.log(`\n📋 2025 시트 M열 전체 데이터: ${M열전체데이터.length}개 고유값`);

    // 열 인덱스를 Excel 열 문자로 변환하는 함수
    function getColumnLetter(index) {
      let result = '';
      index++;
      while (index > 0) {
        index--;
        result = String.fromCharCode(65 + (index % 26)) + result;
        index = Math.floor(index / 26);
      }
      return result;
    }

    // forEach 대신 for...of 루프 사용 (async/await 지원)
    for (let index = 0; index < dataRows2025.length; index++) {
      const row = dataRows2025[index];
      // I열(비고) 값 가져오기
      const 비고2025 = String(row[비고Index2025] || "").trim();
      
      // ============================================
      // 0번째 조건: I열 데이터 변경 여부 확인 (두 번째 실행부터 실행)
      //  - 기존 결과 파일이 있을 때만 실행 (첫 번째 실행 시에는 건너뜀)
      //  - moca 법인은 2024moca 학습 데이터 변경을 항상 반영해야 하므로 0단계 스킵
      // ============================================
      if (!isMocaFile && 기존I열값맵.size > 0 && 기존K열값맵.size > 0) {
        if (기존I열값맵.has(index) && 기존K열값맵.has(index)) {
          const 기존I값 = String(기존I열값맵.get(index) || "").trim();
          const 현재I값 = String(비고2025 || "").trim();
          
          // I열 값이 변경되지 않았으면 기존 K열 값으로 설정하고 다음 행으로
          if (현재I값 === 기존I값) {
            const 기존K값 = 기존K열값맵.get(index);
            row[K열Index2025] = 기존K값;
            continue; // 기존 K열 값 사용, 처리 건너뜀
          }
        }
      }
      
      processedCount++;
      
      if (!비고2025 || 비고2025 === "") {
        row[K열Index2025] = "-";
        noMatchCount++;
        continue;
      }

      // K열 값 추출: 조건부 처리 (우선순위 순서)
      let k열값 = null;
      let 매칭조건 = null;

      // ============================================
      // 첫 번째 우선순위: I열(비고)에서 "월|" 패턴 추출 후 M열에서 정확 일치 (100% 매칭)
      // ============================================
      if (비고2025 && 비고2025.trim() !== "") {
        // "월|" 패턴 찾기
        const 월패턴시작 = 비고2025.indexOf("월|");
        
        if (월패턴시작 !== -1) {
          // "월|" 다음 위치부터 시작
          const 추출시작위치 = 월패턴시작 + 2; // "월|" 길이 = 2
          
          // 다음 "|" 찾기 (추출시작위치 이후)
          const 다음파이프 = 비고2025.indexOf("|", 추출시작위치);
          
          if (다음파이프 !== -1) {
            // "월|" 다음부터 다음 "|" 전까지 텍스트 추출
            const 추출된텍스트 = 비고2025.substring(추출시작위치, 다음파이프).trim();
            
            if (추출된텍스트 && 추출된텍스트 !== "") {
              // M열 전체 데이터에서 정확히 일치하는 값 찾기 (100% 매칭)
              const 정확일치인덱스 = M열전체데이터.findIndex(m값 => 
                String(m값 || "").trim() === 추출된텍스트
              );
              
              if (정확일치인덱스 !== -1) {
                // 정확히 일치하는 값이 있으면 K열에 입력하고 종료 (OpenAI 호출 건너뜀)
                k열값 = M열전체데이터[정확일치인덱스];
                매칭조건 = "첫번째조건";
              }
            }
          }
        }
      }

      // ============================================
      // 두 번째 조건: 첫 번째 조건에서 100% 일치로 추출된 값 없으면 "기타"로 표시
      // ============================================
      // 첫 번째 조건이 실패한 경우에만 두 번째 조건 적용
      if (!k열값 || k열값.trim() === "") {
        k열값 = "기타";
        매칭조건 = "두번째조건";
        
        if (isMocaFile && index < 5) {
          console.log(`   🔍 [moca 디버그] 행 ${index + 1}:`);
          console.log(`      비고: "${비고2025.substring(0, 50)}"`);
          console.log(`      첫 번째 조건 실패 → "기타"로 설정`);
        }
      }

      // 최종 결과 적용
      if (k열값) {
        row[K열Index2025] = k열값;
        matchedCount++;
      } else {
        // 모든 조건 실패 시 "-" 입력
        row[K열Index2025] = "-";
        noMatchCount++;
      }
    }

    console.log("\n📊 처리 결과:");
    console.log(`   - 처리된 행: ${processedCount}개`);
    console.log(`   - 매칭 성공 (K열 채움): ${matchedCount}개`);
    console.log(`   - 매칭 실패: ${noMatchCount}개`);
    // OpenAI 호출 횟수 로그 제거 (OpenAI 사용 안 함)

    // 결과를 새 파일로 저장
    //  - MOCA/기타 법인은 기존 로직 유지 (기존 result 시트 기반으로 다른 열 보존)
    // 첫 번째 실행: 원본 파일의 시트를 기반으로 함
    // 두 번째 실행부터: (MOCA/기타만) 기존 결과 파일의 시트를 기반으로 함
    // 🔥 resultSheetName은 위에서 이미 선언됨 (26번 줄)
    let updatedSheet2025;

    if (isMocaFile) {
      // ✅ MOCA: 항상 원본 2025moca 시트 전체 복사 (원본 변경사항 반영)
      // 원본 시트를 배열로 읽어서 모든 데이터 보존 후 다시 시트로 변환
      // 🔥 data2025는 이미 위에서 읽었으므로 재사용 (중복 읽기 방지)
      const 원본데이터배열 = data2025; // 이미 읽은 데이터 재사용
      console.log(`   📋 MOCA 파일: 원본 ${processingSheetName} 시트 데이터 ${원본데이터배열.length}개 행 읽음`);
      
      // 배열을 다시 시트로 변환 (모든 열 보존)
      updatedSheet2025 = xlsx.utils.aoa_to_sheet(원본데이터배열);
      
      // 🔥 변환된 시트의 범위 확인
      const convertedRange = updatedSheet2025['!ref'] ? xlsx.utils.decode_range(updatedSheet2025['!ref']) : null;
      if (convertedRange) {
        console.log(`   📊 변환된 시트 범위: ${updatedSheet2025['!ref']} (행 ${convertedRange.e.r + 1}까지)`);
      }
      
      console.log(`   ✅ MOCA 파일: 원본 ${processingSheetName} 시트의 모든 열을 보존하여 다시 생성합니다. (원본 변경사항 반영)`);
    } else {
      // ✅ 기존 법인: 항상 원본 시트 전체 복사 (원본 변경사항 반영)
      // 원본 시트를 배열로 읽어서 모든 데이터 보존 후 다시 시트로 변환
      const 원본데이터배열 = xlsx.utils.sheet_to_json(sheet2025, { header: 1, defval: "" });
      // 배열을 다시 시트로 변환 (모든 열 보존)
      updatedSheet2025 = xlsx.utils.aoa_to_sheet(원본데이터배열);
      console.log(`   📋 기존 법인 파일: 원본 ${processingSheetName} 시트의 모든 열을 보존하여 다시 생성합니다. (원본 변경사항 반영)`);
    }
    
    // 🔥 모든 행을 처리하기 위해 dataRows2025.length만큼 반복
    // updatedSheet2025는 이미 원본 데이터 전체를 포함하므로, K열 값만 업데이트하면 됨
    for (let i = 0; i < dataRows2025.length; i++) {
      const row = dataRows2025[i];
      const excelRowNumber = headerRow2025 + 2 + i; // Excel 행 번호 (헤더 + 1행부터 시작)
      const K열문자 = getColumnLetter(K열Index2025);
      const K열셀주소 = `${K열문자}${excelRowNumber}`;
      
        // 0번째 조건: I열 변경 없음인 경우 기존 결과 파일의 K열 값 직접 사용 (변화 없음)
        // 🔥 moca 법인은 2024moca 학습 변경을 반영해야 하므로 0단계 스킵
        if (!isMocaFile && 기존I열값맵.size > 0 && 기존K열값맵.size > 0) {
          if (기존I열값맵.has(i) && 기존K열값맵.has(i)) {
            const 기존I값 = String(기존I열값맵.get(i) || "").trim();
            const 현재I값 = String(row[비고Index2025] || "").trim();
            
            // I열 값이 변경되지 않았으면 기존 결과 파일의 K열 값 직접 사용 (변화 없음)
            if (현재I값 === 기존I값) {
              const 기존K값 = 기존K열값맵.get(i);
              if (기존K값 !== undefined && 기존K값 !== null && 기존K값 !== "") {
                // 기존 결과 파일의 K열 값을 그대로 유지 (변화 없음)
                // updatedSheet2025는 이미 기존 결과 파일의 시트를 복사했으므로 K열 값은 이미 있음
                // 따라서 아무것도 하지 않음 (기존 값 유지)
                continue; // 다음 행으로 (변화 없음)
              }
            }
          }
        }
      
      // I열 변경됨 또는 새로 처리된 행: row[K열Index2025]에서 새로운 값 가져오기
      const k열값 = String(row[K열Index2025] || "").trim();
      
      if (isMocaFile && i < 5) {
        console.log(`   📝 [moca 디버그] 행 ${i + 1} (Excel 행 ${excelRowNumber}): K열값="${k열값}", 셀주소="${K열셀주소}"`);
      }
      
      // 🔥 K열 값이 있으면 무조건 업데이트 (셀이 없으면 생성)
      if (k열값 && k열값 !== "") {
        // 새로운 값이 있으면 업데이트
        if (!updatedSheet2025[K열셀주소]) {
          updatedSheet2025[K열셀주소] = {};
        }
        updatedSheet2025[K열셀주소].v = k열값;
        updatedSheet2025[K열셀주소].t = 's'; // 텍스트 타입
        delete updatedSheet2025[K열셀주소].f; // 수식 제거
        
        if (isMocaFile && i < 5) {
          console.log(`   ✅ [moca 디버그] 행 ${i + 1}: 시트에 값 저장됨 "${k열값}"`);
        }
      } else if (기존K열값맵.has(i)) {
        // 새로운 값이 없으면 기존 결과 파일의 K열 값 유지
        const 기존K값 = 기존K열값맵.get(i);
        if (기존K값 !== undefined && 기존K값 !== null && 기존K값 !== "") {
          if (!updatedSheet2025[K열셀주소]) {
            updatedSheet2025[K열셀주소] = {};
          }
          updatedSheet2025[K열셀주소].v = 기존K값;
          updatedSheet2025[K열셀주소].t = 's'; // 텍스트 타입
          delete updatedSheet2025[K열셀주소].f; // 수식 제거
        }
      }
      // k열값이 없고 기존K열값맵에도 없으면 아무것도 하지 않음 (기존 값 유지)
    }
    
    // 🔥 처리 완료 후 최종 시트 범위 확인
    if (isMocaFile) {
      const finalRange = updatedSheet2025['!ref'] ? xlsx.utils.decode_range(updatedSheet2025['!ref']) : null;
      if (finalRange) {
        console.log(`   📊 최종 시트 범위: ${updatedSheet2025['!ref']} (행 ${finalRange.e.r + 1}까지)`);
        console.log(`   ✅ 처리 완료: ${dataRows2025.length}개 행 처리됨`);
      }
    }

    const newWorkbook = xlsx.utils.book_new();
    newWorkbook.SheetNames = workbook.SheetNames;
    newWorkbook.Sheets = { ...workbook.Sheets };
    newWorkbook.Sheets[resultSheetName] = updatedSheet2025; // 🔥 moca 파일은 2025moca 시트로 저장

    try {
      xlsx.writeFile(newWorkbook, outputFile);
      
      // 🔥 최종 결과 파일의 행 수 확인
      if (isMocaFile) {
        const finalWorkbook = xlsx.readFile(outputFile);
        const finalSheet = finalWorkbook.Sheets[resultSheetName];
        if (finalSheet && finalSheet['!ref']) {
          const finalRange = xlsx.utils.decode_range(finalSheet['!ref']);
          console.log(`\n✅ 결과 파일 생성 완료: ${outputFile}`);
          console.log(`   📊 결과 파일 행 수: ${finalRange.e.r + 1}개 행`);
          console.log(`   → 원본 파일(${inputFile})은 수정하지 않았습니다.`);
        } else {
          console.log(`\n✅ 결과 파일 생성 완료: ${outputFile}`);
          console.log(`   → 원본 파일(${inputFile})은 수정하지 않았습니다.`);
        }
      } else {
        console.log(`\n✅ 결과 파일 생성 완료: ${outputFile}`);
        console.log(`   → 원본 파일(${inputFile})은 수정하지 않았습니다.`);
      }
      
      return {
        success: true,
        processed: processedCount,
        matched: matchedCount,
        noMatch: noMatchCount,
        outputFile: outputFile
      };
    } catch (error) {
      if (error.code === 'EBUSY' || error.code === 'EACCES') {
        throw new Error("❌ 오류: 파일이 다른 프로그램에서 열려있습니다! Excel에서 파일을 닫고 다시 실행해주세요.");
      } else {
        throw error;
      }
    }
  } catch (error) {
    console.error("❌ 처리 중 오류 발생:", error.message);
    throw error;
  }
}

