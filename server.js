import 'dotenv/config';

import express from "express";
import cors from "cors";
import compression from "compression";
import { getExcelData, getSheetNames } from "./dataService.js";
import dotenv from "dotenv";
// import OpenAI from "openai"; // OpenAI 기능 제거됨
import path from "path";
import { fileURLToPath } from "url";
import fs from "fs";
import xlsx from "xlsx";
import os from "os";
import { exec } from "child_process";
import { processExcelFile } from "./processExcel.mjs";
// SQL 연동을 위한 패키지
import sql from 'mssql';  // SQL Server 사용

dotenv.config();

const app = express();
// 🔥 포트는 환경 변수 PORT를 우선 사용, 없으면 3000
const PORT = process.env.PORT ? Number(process.env.PORT) || 3000 : 3000;
const REDUCE_LOG = process.env.REDUCE_LOG === 'true';
const SKIP_FILE_WRITE = process.env.SKIP_FILE_WRITE === 'true';
const CACHE_TTL_MS = (Number(process.env.RESPONSE_CACHE_TTL_MS) || 5 * 60 * 1000); // 기본 5분
const responseCache = new Map();

// 필요 시 로그 최소화 (info/debug 수준만)
if (REDUCE_LOG) {
  const noop = () => {};
  console.log = noop;
  console.debug = noop;
}

// __dirname 설정
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// CORS와 JSON 파서는 먼저 설정
// 모든 origin 허용 (개발 및 네트워크 공유용)
app.use(cors({
    origin: '*', // 모든 origin 허용
    methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'Authorization'],
    credentials: false
}));
// 응답 압축 (네트워크 전송량 감소)
app.use(compression());
app.use(express.json());

// 🔥 모든 요청 로깅 미들웨어 (디버깅용)
app.use((req, res, next) => {
  const timestamp = new Date().toLocaleTimeString('ko-KR', { hour12: false });
  console.log(`\n🌐 [${timestamp}] ${req.method} ${req.path}`);
  console.log(`   📍 요청 URL: ${req.protocol}://${req.get('host')}${req.originalUrl}`);
  if (Object.keys(req.query).length > 0) {
    console.log(`   📋 쿼리 파라미터:`, req.query);
  }
  if (req.body && Object.keys(req.body).length > 0) {
    console.log(`   📦 요청 본문:`, req.body);
  }
  next();
});

// 엑셀 파일 경로 설정
// 환경 변수에서 가져오거나 기본값 사용
// 🔥 기존법인 파일은 더 이상 사용하지 않음 (MOCA 파일만 사용)
// const EXCEL_FILE_PATH = "./match_data_all.xlsx"; // 기존법인 파일 - 사용 안 함
const EXCEL_SHEET_NAME = process.env.EXCEL_SHEET_NAME || "2025";

// 🔥 추가 법인 엑셀 파일 경로 설정
// 여러 법인의 데이터를 병합하기 위한 파일 경로 배열
// 환경 변수 ADDITIONAL_EXCEL_FILES에 쉼표로 구분하여 추가 가능
// 예: ADDITIONAL_EXCEL_FILES="./match_data_moca.xlsx,./match_data_other.xlsx"
const ADDITIONAL_EXCEL_FILES = (
  process.env.ADDITIONAL_EXCEL_FILES 
  ? process.env.ADDITIONAL_EXCEL_FILES.split(',').map(f => f.trim()).filter(f => f)
    : ["./match_data_moca.xlsx"]
);

// 🔥 미정산 전용 파일 경로 설정
// 미정산 데이터만 있는 별도 파일들
const UNSETTLED_EXCEL_FILES = process.env.UNSETTLED_EXCEL_FILES 
  ? process.env.UNSETTLED_EXCEL_FILES.split(',').map(f => f.trim()).filter(f => f)
  : ["./match_data_미결_moca.xlsx"]; // 🔥 match_data_미결_moca.xlsx 기본 추가

// 절대 경로로 변환 (상대 경로인 경우)
function getExcelFilePath(filePath) {
  if (path.isAbsolute(filePath)) {
    return filePath;
  }
  return path.resolve(__dirname, filePath);
}

// 정산월 보정: '2025-06' 이외 형태도 강제로 텍스트로 처리
function normalizeSettlementMonth(value) {
  if (value === undefined || value === null) return null;
  const raw = String(value).trim();
  if (!raw) return null;

  // 이미 YYYY-MM 형식인 경우 그대로 반환
  const yyyyMMMatch = raw.match(/^(\d{4})-(\d{2})$/);
  if (yyyyMMMatch) {
    const year = yyyyMMMatch[1];
    const month = yyyyMMMatch[2];
    const monthNum = Number(month);
    if (monthNum >= 1 && monthNum <= 12) {
      return `${year}-${month}`;
    }
  }

  const digitsOnly = raw.replace(/[^0-9]/g, "");
  if (digitsOnly.length >= 6) {
    const year = digitsOnly.slice(0, 4);
    const month = digitsOnly.slice(4, 6);
    const monthNum = Number(month);
    if (monthNum >= 1 && monthNum <= 12) {
      return `${year}-${month.padStart(2, "0")}`;
    }
  }

  const match = raw.match(/(\d{4}).*?(\d{1,2})/);
  if (match) {
    const year = match[1];
    const month = match[2].padStart(2, "0");
    const monthNum = Number(month);
    if (monthNum >= 1 && monthNum <= 12) {
      return `${year}-${month}`;
    }
  }

  return null;
}

let responseData = null;

// ===================================================
// 📌 미정산 상세내역 계정명 계산 함수
// match_data_AI.xlsx 파일을 사용하여 계정명 계산
// ===================================================
let 학습데이터캐시 = null;
let C열합계잔액시산표계정명목록캐시 = null; // match_data_AI.xlsx의 2024 시트 C열 데이터
let 적요목록2024캐시 = null;
let 정규화된적요목록2024캐시 = null;

// 🔥 SQL 미정산 데이터의 계정명 계산 결과 캐시 (비고값 -> 계정명 매핑)
// 비고값이 변경되거나 새로운 데이터가 추가될 때만 OpenAI 재실행
const unsettledAccountNameCache = new Map(); // key: 비고값 (정규화), value: { 계정명, 매칭방법, 매치율 }

// match_data_AI.xlsx에서 학습 데이터 로드
async function loadLearningDataFromMatchDataAI() {
  // 텍스트 정규화 함수 (캐시에서도 사용)
  function removeDates(text) {
    if (!text) return '';
    return String(text)
      .replace(/\d{2,4}년\s*\d{1,2}월/g, '')
      .replace(/\d{2,4}\.\d{1,2}/g, '')
      .replace(/\d{4}-\d{2}-\d{2}/g, '')
      .replace(/\d{8}/g, '')
      .replace(/\d{4}년/g, '')
      .replace(/\d{1,2}월/g, '');
  }

  function normalizeText(text) {
    if (!text) return '';
    let normalized = String(text);
    normalized = removeDates(normalized);
    normalized = normalized.replace(/\s+/g, '');
    // 🔥 번호 패턴 제거: (1), (2), (3) 등 제거하여 동일한 적요를 통일
    normalized = normalized.replace(/\(\d+\)/g, '');
    normalized = normalized.replace(/[^\w가-힣]/g, '');
    normalized = normalized.toLowerCase();
    return normalized;
  }

  if (학습데이터캐시 && C열합계잔액시산표계정명목록캐시) {
    return {
      학습데이터: 학습데이터캐시,
      C열합계잔액시산표계정명목록: C열합계잔액시산표계정명목록캐시, // match_data_AI.xlsx의 2024 시트 C열 데이터
      적요목록2024: 적요목록2024캐시,
      정규화된적요목록2024: 정규화된적요목록2024캐시,
      normalizeText  // normalizeText 함수도 반환
    };
  }

  try {
    const matchDataAIPath = path.join(__dirname, 'match_data_AI.xlsx');
    if (!fs.existsSync(matchDataAIPath)) {
      console.log('⚠️ match_data_AI.xlsx 파일이 없습니다. 계정명 계산을 건너뜁니다.');
      return null;
    }

    const workbook = xlsx.readFile(matchDataAIPath);
    const sheet2024 = workbook.Sheets['2024'];
    
    if (!sheet2024) {
      console.log('⚠️ match_data_AI.xlsx에서 2024 시트를 찾을 수 없습니다.');
      return null;
    }

    // 2024 시트 데이터 파싱
    const data2024 = xlsx.utils.sheet_to_json(sheet2024, { header: 1, defval: '' });
    
    // 헤더 행 찾기
    let headerRow2024 = 0;
    for (let i = 0; i < Math.min(10, data2024.length); i++) {
      const row = data2024[i] || [];
      const firstCell = String(row[0] || '').trim();
      if (firstCell.includes('적요') || firstCell.includes('계정명')) {
        headerRow2024 = i;
        break;
      }
    }

    const header2024 = data2024[headerRow2024] || [];
    
    // A열(인덱스 0): 적요, B열(인덱스 1): 계정명, C열(인덱스 2): 합계잔액시산표 계정명
    const 적요Index2024 = header2024.findIndex(h => String(h || '').includes('적요')) !== -1 
      ? header2024.findIndex(h => String(h || '').includes('적요'))
      : 0; // 기본값: A열
    const 계정명Index2024 = header2024.findIndex(h => String(h || '').includes('계정명')) !== -1
      ? header2024.findIndex(h => String(h || '').includes('계정명'))
      : 1; // 기본값: B열
    const 합계잔액시산표계정명Index2024 = header2024.findIndex(h => 
      String(h || '').includes('합계잔액시산표')
    ) !== -1
      ? header2024.findIndex(h => String(h || '').includes('합계잔액시산표'))
      : 2; // 기본값: C열

    // 학습 데이터 생성 (A열: 적요, B열: 계정명, C열: 합계잔액시산표 계정명)
    const dataRows2024 = data2024.slice(headerRow2024 + 1);
    const 학습데이터 = [];
    dataRows2024.forEach((row) => {
      const 적요 = String(row[적요Index2024] !== -1 ? row[적요Index2024] : row[0] || '').trim();
      const 계정명 = String(row[계정명Index2024] !== -1 ? row[계정명Index2024] : row[1] || '').trim();
      const 합계잔액시산표계정명 = String(
        row[합계잔액시산표계정명Index2024] !== -1 
          ? row[합계잔액시산표계정명Index2024] 
          : row[2] || ''
      ).trim();
      if (적요 && 계정명) {
        학습데이터.push({ 
          적요, 
          계정명,
          합계잔액시산표계정명: 합계잔액시산표계정명 || 계정명
        });
      }
    });

    // C열(합계잔액시산표 계정명) 목록 가져오기 (2024 시트에서)
    // C열은 인덱스 2 (0-based, A=0, B=1, C=2)
    const C열Index = 합계잔액시산표계정명Index2024 !== -1 ? 합계잔액시산표계정명Index2024 : 2;
    const C열합계잔액시산표계정명목록 = []; // match_data_AI.xlsx의 2024 시트 C열 데이터
    dataRows2024.forEach(row => {
      const c값 = String(row[C열Index] || '').trim();
      if (c값 && c값 !== '' && c값 !== '-' && !C열합계잔액시산표계정명목록.includes(c값)) {
        C열합계잔액시산표계정명목록.push(c값);
      }
    });

    // 적요 목록 생성
    const 적요목록2024 = 학습데이터.map(d => d.적요);
    
    // 텍스트 정규화 함수
    function removeDates(text) {
      if (!text) return '';
      return String(text)
        .replace(/\d{2,4}년\s*\d{1,2}월/g, '')
        .replace(/\d{2,4}\.\d{1,2}/g, '')
        .replace(/\d{4}-\d{2}-\d{2}/g, '')
        .replace(/\d{8}/g, '')
        .replace(/\d{4}년/g, '')
        .replace(/\d{1,2}월/g, '');
    }

    function normalizeText(text) {
      if (!text) return '';
      let normalized = String(text);
      normalized = removeDates(normalized);
      normalized = normalized.replace(/\s+/g, ''); // 띄어쓰기 제거
      // 🔥 번호 패턴 제거: (1), (2), (3) 등 제거하여 동일한 적요를 통일
      normalized = normalized.replace(/\(\d+\)/g, '');
      normalized = normalized.replace(/[^\w가-힣]/g, ''); // 특수 문자 제거
      normalized = normalized.toLowerCase(); // 소문자 변환
      return normalized;
    }

    const 정규화된적요목록2024 = 적요목록2024.map(적요 => normalizeText(적요));

    // 캐시에 저장
    학습데이터캐시 = 학습데이터;
    C열합계잔액시산표계정명목록캐시 = C열합계잔액시산표계정명목록;
    적요목록2024캐시 = 적요목록2024;
    정규화된적요목록2024캐시 = 정규화된적요목록2024;

    console.log(`📚 match_data_AI.xlsx 학습 데이터 로드 완료: ${학습데이터.length}개 행, C열(합계잔액시산표 계정명) 데이터: ${C열합계잔액시산표계정명목록.length}개`);

    return {
      학습데이터,
      C열합계잔액시산표계정명목록, // match_data_AI.xlsx의 2024 시트 C열 데이터
      적요목록2024,
      정규화된적요목록2024,
      normalizeText
    };
  } catch (error) {
    console.error('❌ match_data_AI.xlsx 로드 오류:', error);
    return null;
  }
}

// match_data_moca.xlsx에서 M열(합계잔액시산표 계정명) 목록 로드
let M열합계잔액시산표계정명목록캐시 = null; // match_data_moca.xlsx의 2025moca 시트 M열 데이터

async function loadMColumnFromMatchDataMoca() {
  // 캐시가 있으면 재사용
  if (M열합계잔액시산표계정명목록캐시) {
    return M열합계잔액시산표계정명목록캐시;
  }

  try {
    const mocaFilePath = path.join(__dirname, 'match_data_moca.xlsx');
    if (!fs.existsSync(mocaFilePath)) {
      console.log('⚠️ match_data_moca.xlsx 파일이 없습니다.');
      return [];
    }

    const workbook = xlsx.readFile(mocaFilePath);
    const sheet2025moca = workbook.Sheets['2025moca'];
    
    if (!sheet2025moca) {
      console.log('⚠️ match_data_moca.xlsx에서 2025moca 시트를 찾을 수 없습니다.');
      return [];
    }

    // 2025moca 시트 데이터 파싱
    const data2025moca = xlsx.utils.sheet_to_json(sheet2025moca, { header: 1, defval: '' });
    
    // 헤더 행 찾기
    let headerRow2025moca = 0;
    for (let i = 0; i < Math.min(10, data2025moca.length); i++) {
      const row = data2025moca[i] || [];
      const firstCell = String(row[0] || '').trim();
      if (firstCell.includes('비고') || firstCell.includes('적요') || firstCell.includes('전표번호')) {
        headerRow2025moca = i;
        break;
      }
    }

    const header2025moca = data2025moca[headerRow2025moca] || [];
    
    // M열(합계잔액시산표 계정명) 인덱스 찾기
    // M열은 인덱스 12 (0-based, A=0, B=1, ..., M=12)
    let M열Index2025moca = header2025moca.findIndex(h => 
      String(h || '').includes('합계잔액시산표')
    );
    
    if (M열Index2025moca === -1) {
      M열Index2025moca = 12; // 기본값: M열 (인덱스 12)
    }

    // M열(합계잔액시산표 계정명) 목록 가져오기
    const dataRows2025moca = data2025moca.slice(headerRow2025moca + 1);
    const M열합계잔액시산표계정명목록 = [];
    dataRows2025moca.forEach(row => {
      const m값 = String(row[M열Index2025moca] || '').trim();
      if (m값 && m값 !== '' && m값 !== '-' && !M열합계잔액시산표계정명목록.includes(m값)) {
        M열합계잔액시산표계정명목록.push(m값);
      }
    });

    // 캐시에 저장
    M열합계잔액시산표계정명목록캐시 = M열합계잔액시산표계정명목록;

    console.log(`📚 match_data_moca.xlsx 2025moca 시트 M열(합계잔액시산표 계정명) 데이터 로드 완료: ${M열합계잔액시산표계정명목록.length}개`);

    return M열합계잔액시산표계정명목록;
  } catch (error) {
    console.error('❌ match_data_moca.xlsx M열 로드 오류:', error);
    return [];
  }
}

// ⭐ SQL 비고에서 계정명 추출 함수 추가
function extractAccountNameFromSQL(note) {
  if (!note) return "";
  const parts = note.split("|");
  if (parts.length < 2) return "";
  return parts[1].trim();   // 계정명 100% 추출
}

// 비고에서 "월|" 패턴 추출
// 숫자 + '월|' 패턴에서 계정명 추출 후 M열 데이터와 100% 일치 비교
function extractAccountNameFromNote(비고값, M열합계잔액시산표계정명목록) {
  if (!비고값 || !M열합계잔액시산표계정명목록 || M열합계잔액시산표계정명목록.length === 0) {
    return null;
  }

  // 숫자 + '월|' 패턴 찾기 (예: 10월|, 11월|, 25년11월| 등)
  // 패턴: 숫자+월| 다음부터 다음 | 전까지 텍스트 추출
  const match = 비고값.match(/\d+월\|(.+?)\|/);
  
  if (match && match[1]) {
    const 추출된계정명 = match[1].trim();
    
    if (추출된계정명 && 추출된계정명 !== '') {
      // M열(합계잔액시산표 계정명) 목록에서 정확히 일치하는 값 찾기 (100% 매칭)
      const 정확일치인덱스 = M열합계잔액시산표계정명목록.findIndex(m값 => 
        String(m값 || '').trim() === 추출된계정명
      );
        
      if (정확일치인덱스 !== -1) {
        return M열합계잔액시산표계정명목록[정확일치인덱스];
      }
    }
  }
  
  return null;
}


// 미정산 상세내역 계정명 계산 (메인 함수)
async function calculateUnsettledAccountName(비고값, returnDetail = false, useCache = true) {
  try {
    // 🔥 캐시 확인
    if (useCache && unsettledAccountNameCache.has(비고값.trim())) {
      const 캐시된값 = unsettledAccountNameCache.get(비고값.trim());
      // 캐시된 값이 "기타"이면 무시하고 재계산 (C열 목록이 업데이트되었을 수 있음)
      if (캐시된값 && typeof 캐시된값 === 'object' && 캐시된값.계정명 === '기타') {
        unsettledAccountNameCache.delete(비고값.trim());
      } else if (typeof 캐시된값 === 'string' && 캐시된값 === '기타') {
        unsettledAccountNameCache.delete(비고값.trim());
      } else {
        const result = returnDetail ? 캐시된값 : (typeof 캐시된값 === 'object' ? 캐시된값.계정명 : 캐시된값);
        return result;
      }
    }

    // 1번 조건: 비고에서 "월|" 패턴 추출 후 M열 데이터와 100% 일치 비교
    // M열 목록 로드 (1번 조건용) - match_data_moca 파일의 M열(합계잔액시산표 계정명)
    const M열합계잔액시산표계정명목록 = await loadMColumnFromMatchDataMoca();
    
    if (M열합계잔액시산표계정명목록 && M열합계잔액시산표계정명목록.length > 0) {
      // 1번 조건: "월|" 패턴 추출 후 M열 목록과 비교
      const extractedAccountName = extractAccountNameFromNote(비고값, M열합계잔액시산표계정명목록);
      if (extractedAccountName) {
        const result = returnDetail ? { 계정명: extractedAccountName, 매칭방법: '월|패턴추출', 매치율: 1.0 } : extractedAccountName;
        // 캐시에 저장
        if (useCache) {
          unsettledAccountNameCache.set(비고값.trim(), returnDetail ? result : { 계정명: result, 매칭방법: '월|패턴추출', 매치율: 1.0 });
        }
        return result;
      }
    }

    // 2번 조건: 첫 번째 조건에서 100% 매치율 안나오는 계정명 "기타"로 표기
    const result = returnDetail ? { 
      계정명: '기타', 
      매칭방법: '매칭실패', 
      매치율: 0 
    } : '기타';
    // 캐시에 저장
    if (useCache) {
      unsettledAccountNameCache.set(비고값.trim(), returnDetail ? result : { 계정명: result, 매칭방법: '매칭실패', 매치율: 0 });
    }
    return result;
  } catch (error) {
    console.error(`   ❌ calculateUnsettledAccountName 내부 오류:`, error);
    console.error(`   - 오류 스택:`, error.stack);
    const errorResult = returnDetail ? { 
      계정명: '기타', 
      매칭방법: '매칭오류', 
      매치율: 0 
    } : '기타';
    if (useCache) {
      unsettledAccountNameCache.set(비고값.trim(), returnDetail ? errorResult : { 계정명: errorResult, 매칭방법: '매칭오류', 매치율: 0 });
    }
    console.log(`${"=".repeat(60)}\n`);
    return errorResult;
  }
}

// ===================================================
// 📌 SQL 데이터 조회 함수 (공통)
// 지급일 기준으로 데이터를 SQL에서 가져옴
// type: 'settled' (정산) 또는 'unsettled' (미정산)
// period: 조회 기간 (예: "2025-01 ~ 2025-12")
// ===================================================
async function getSettlementDataFromSQL(userName = null, type = 'settled', period = null) {
  try {
    // SQL 연결 정보 확인
    const dbConfig = {
      server: process.env.DB_HOST || process.env.DB_SERVER || 'localhost',
      port: parseInt(process.env.DB_PORT || '1433'), // SQL Server 기본 포트
      user: process.env.DB_USER,
      password: process.env.DB_PASSWORD,
      database: process.env.DB_NAME || process.env.DB_DATABASE,
      options: {
        encrypt: process.env.DB_ENCRYPT === 'true', // Azure SQL 사용 시 true
        trustServerCertificate: process.env.DB_TRUST_CERT === 'true' || true, // 개발 환경에서 인증서 검증 건너뛰기
        enableArithAbort: true
      }
    };

    // 타입에 따라 테이블 선택
    let tableName = '';
    if (type === 'settled') {
      // 정산 상세내역: [dbo].[ERP_이체내역조회]
      tableName = process.env.DB_TABLE_SETTLED || '[dbo].[ERP_이체내역조회]';
    } else {
      // 미정산 상세내역: [dbo].[ERP_전표상세조회_자금]
      tableName = process.env.DB_TABLE_UNSETTLED || '[dbo].[ERP_전표상세조회_자금]';
    }

    // 환경 변수가 설정되지 않았으면 빈 배열 반환
    if (!dbConfig.user || !dbConfig.password || !dbConfig.database) {
      console.log(`\n${"=".repeat(80)}`);
      console.log('⚠️ SQL 연결 정보가 설정되지 않았습니다.');
      console.log(`   타입: ${type === 'settled' ? '정산' : '미정산'}`);
      console.log(`   테이블: ${tableName}`);
      console.log(`   사용자 필터: ${userName || '없음 (전체)'}`);
      console.log('\n   필요한 환경 변수:');
      console.log('   - DB_HOST 또는 DB_SERVER');
      console.log('   - DB_PORT (기본값: 1433)');
      console.log('   - DB_USER');
      console.log('   - DB_PASSWORD');
      console.log('   - DB_NAME 또는 DB_DATABASE');
      console.log('   - DB_TABLE_SETTLED (정산)');
      console.log('   - DB_TABLE_UNSETTLED (미정산)');
      console.log(`\n   현재 설정값:`);
      console.log(`   - DB_HOST: ${process.env.DB_HOST || process.env.DB_SERVER || '없음'}`);
      console.log(`   - DB_PORT: ${process.env.DB_PORT || '1433 (기본값)'}`);
      console.log(`   - DB_USER: ${process.env.DB_USER ? '설정됨' : '없음'}`);
      console.log(`   - DB_PASSWORD: ${process.env.DB_PASSWORD ? '설정됨' : '없음'}`);
      console.log(`   - DB_NAME: ${process.env.DB_NAME || process.env.DB_DATABASE || '없음'}`);
      console.log(`${"=".repeat(80)}\n`);
      return [];
    }

    const typeLabel = type === 'settled' ? '정산' : '미정산';
    const dateCondition = period ? `(지급일 기준: ${period})` : (type === 'settled' ? '(2025-11 이후)' : '(모든 미정산 데이터)');
    console.log(`📊 SQL Server에서 ${typeLabel} 데이터 조회 시작 ${dateCondition}...`);
    console.log(`   서버: ${dbConfig.server}:${dbConfig.port}`);
    console.log(`   데이터베이스: ${dbConfig.database}`);
    console.log(`   테이블: ${tableName}`);
    console.log(`   사용자 필터: ${userName || '없음 (전체)'}`);
    console.log(`   조회 기간: ${period || '없음'}`);

    // SQL Server 쿼리 생성
    let query = '';
    
    if (type === 'unsettled') {
      // 미정산 상세내역: [dbo].[ERP_전표상세조회_자금]
      // 순서:
      // 1. 사용자 컬럼으로 필터링 (userName이 있으면)
      //    - 사용자 조회 시: 반제일 IS NULL AND 사용자 LIKE 조건
      //    - 사용자 없을 때: 모든 데이터 조회
      // 3. 만약 계정명이 '미지급금_사내' 있으면 사용처에 사용자명 넣기 (데이터 변환 단계에서 처리)
      // 4. 정산월, 사용처, 비고, 사용금액 등등 컬럼에 맞게 데이터 넣기
      query = `
        SELECT 
          정산월 AS settlementMonth,
          만기일 AS paymentDate,
          사용처 AS merchant,
          사용자 AS userColumn,  -- 사용자 컬럼도 가져와서 나중에 사용처에 넣을 수 있도록
          사용금액 AS amount,
          비고 AS note
        FROM ${tableName}
      `;
      
      // 1. 사용자 컬럼으로 필터링 (userName이 있으면)
      // 사용자 조회 시 반제일이 NULL인 데이터만 가져오기
      if (userName) {
        query += ` WHERE 반제일 IS NULL AND 사용자 LIKE @userName`;
      }
      
      // 3번은 데이터 변환 단계에서 처리 (계정명이 '미지급금_사내'인 경우 사용처에 사용자명 넣기)
      
      query += ` ORDER BY 정산월 DESC, 만기일 DESC`;
    } else {
      // 정산 상세내역: [dbo].[ERP_이체내역조회]
      // 지급일(반제일) 기준으로 조회
      query = `
        SELECT 
          정산월 AS settlementMonth,
          반제일 AS paymentDate,
          사용처 AS merchant,
          출금액 AS amount,
          비고 AS note,
          거래처명 AS 거래처명
        FROM ${tableName}
        WHERE 1=1
      `;
      
      // 🔥 SQL 데이터는 항상 2025-11 이후만 조회 (2025-01~2025-10은 엑셀에서 가져옴)
      // 조회 기간이 있어도 정산월 >= '2025-11' 조건은 항상 적용
      query += ` AND 정산월 >= '2025-11'`;
      
      // 조회 기간이 있으면 지급일(반제일) 기준으로 추가 필터링
      // 🔥 조회기간을 1개월 앞당겨서 지급일 기준으로 필터링
      // 예: 조회기간 2025-01~2025-12 → 지급일 2024-12~2025-11 (정산월 2025-01의 지급일이 2024-12일 수 있음)
      // 🔥 단, 조회 기간의 종료 월이 2025-11 이상일 때만 적용 (정산월 >= '2025-11' 조건과 충돌 방지)
      if (period) {
        // period 파싱: "2025-01 ~ 2025-12" 형식
        const periodMatch = period.match(/(\d{4})-(\d{2})\s*~\s*(\d{4})-(\d{2})/);
        if (periodMatch) {
          const [, startYear, startMonth, endYear, endMonth] = periodMatch;
          const endMonthKey = `${endYear}-${endMonth}`;
          
          // 조회 기간의 종료 월이 2025-11 이상일 때만 지급일 필터 적용
          if (endMonthKey >= '2025-11') {
            // 🔥 조회기간을 1개월 앞당김 (예: 2025-01 → 2024-12)
            let adjustedStartYear = parseInt(startYear);
            let adjustedStartMonth = parseInt(startMonth) - 1;
            if (adjustedStartMonth < 1) {
              adjustedStartMonth = 12;
              adjustedStartYear -= 1;
            }
            
            let adjustedEndYear = parseInt(endYear);
            let adjustedEndMonth = parseInt(endMonth) - 1;
            if (adjustedEndMonth < 1) {
              adjustedEndMonth = 12;
              adjustedEndYear -= 1;
            }
            
            const startDate = `${adjustedStartYear}-${String(adjustedStartMonth).padStart(2, '0')}-01`;
            // 마지막 날짜 계산 (예: 2025-11 -> 2025-11-30)
            const lastDay = new Date(adjustedEndYear, adjustedEndMonth, 0).getDate();
            const endDate = `${adjustedEndYear}-${String(adjustedEndMonth).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
            
            // 🔥 정산월 >= '2025-11' 조건과 일치하도록 시작일도 2025-11-01 이상으로 조정
            const finalStartDate = startDate < '2025-11-01' ? '2025-11-01' : startDate;
            
            query += ` AND 반제일 >= '${finalStartDate}' AND 반제일 <= '${endDate}'`;
            console.log(`   📅 지급일 필터 (1개월 앞당김): ${finalStartDate} ~ ${endDate} (조회기간: ${startYear}-${startMonth} ~ ${endYear}-${endMonth})`);
          } else {
            console.log(`   ⚠️ 조회 기간 종료 월(${endMonthKey})이 2025-11 미만이므로 지급일 필터를 적용하지 않습니다. (정산월 >= '2025-11' 조건과 충돌 방지)`);
          }
        }
      }
      
      console.log(`   🔥 SQL 정산월 필터: 정산월 >= '2025-11' (2025-01~2025-10은 엑셀에서 가져옴)`);
      
      if (userName) {
        query += ` AND 거래처명 LIKE @merchant`;
      }
      
      query += ` ORDER BY 정산월 DESC, 반제일 DESC`;
    }

    console.log(`📋 SQL 쿼리: ${query}`);
    if (userName) {
      console.log(`📋 사용자 필터: ${userName}`);
    }

    // 🔥 SQL Server 연결 및 쿼리 실행
    let pool;
    try {
      console.log(`\n${"=".repeat(80)}`);
      console.log(`🔌 SQL Server 연결 시도 중...`);
      console.log(`   서버: ${dbConfig.server}:${dbConfig.port}`);
      console.log(`   데이터베이스: ${dbConfig.database}`);
      console.log(`   사용자: ${dbConfig.user || '없음'}`);
      console.log(`   테이블: ${tableName}`);
      console.log(`${"=".repeat(80)}`);
      
      pool = await sql.connect(dbConfig);
      console.log('✅ SQL Server 연결 성공');
      console.log(`   🔍 연결 정보:`);
      console.log(`      - 서버: ${dbConfig.server}:${dbConfig.port}`);
      console.log(`      - 데이터베이스: ${dbConfig.database}`);
      console.log(`      - 사용자: ${dbConfig.user || '없음'}`);
      console.log(`      - 테이블: ${tableName}`);
      console.log(`      - 타입: ${type === 'settled' ? '정산' : '미정산'}`);
      console.log(`      - 사용자 필터: ${userName || '없음 (전체 조회)'}`);

      const request = pool.request();
      
      // 사용자 필터링 파라미터 추가
      if (userName) {
        if (type === 'unsettled') {
          // 미정산: "사용자" 컬럼으로 필터링
          const filterValue = `%${userName}%`;
          request.input('userName', sql.VarChar, filterValue);
          console.log(`   📋 사용자 필터 파라미터 추가: "${filterValue}"`);
        } else {
          // 정산: "거래처명" 컬럼으로 필터링
          const filterValue = `%${userName}%`;
          request.input('merchant', sql.VarChar, filterValue);
          console.log(`   📋 거래처명 필터 파라미터 추가: "${filterValue}"`);
        }
      }
      
      console.log(`\n📋 SQL 쿼리 실행 중...`);
      console.log(`   전체 쿼리: ${query}`);
      if (userName) {
        console.log(`   사용자 필터: "${userName}"`);
        console.log(`   필터 파라미터: @merchant = "%${userName}%"`);
      }
      
      const result = await request.query(query);
      const rows = result.recordset;
      
      console.log(`\n📊 SQL 쿼리 실행 결과:`);
      console.log(`   조회된 행 수: ${rows.length}개`);
      
      // 🔥 2025-11 데이터의 사용처 값 확인 (디버깅)
      if (type === 'settled' && rows.length > 0) {
        const rows2025_11 = rows.filter(row => {
          const month = row.settlementMonth || row.정산월 || '';
          return month && String(month).startsWith('2025-11');
        });
        if (rows2025_11.length > 0) {
          console.log(`\n🔍 [SQL 쿼리 결과] 2025-11 데이터 사용처 확인:`);
          rows2025_11.slice(0, 3).forEach((row, idx) => {
            console.log(`   ${idx + 1}. 정산월: "${row.settlementMonth || row.정산월}"`);
            console.log(`      SQL 원본 row.merchant: "${row.merchant || '(없음)'}" (타입: ${typeof row.merchant})`);
            console.log(`      SQL 원본 row.사용처: "${row.사용처 || '(없음)'}" (타입: ${typeof row.사용처})`);
            console.log(`      SQL 원본 row.거래처명: "${row.거래처명 || '(없음)'}"`);
            console.log(`      row 객체의 모든 키:`, Object.keys(row).join(', '));
          });
        }
      }
      if (type === 'settled' && userName) {
        console.log(`   🔍 거래처명 필터: "${userName}"`);
        console.log(`   🔍 SQL 쿼리 조건: 정산월 >= '2025-11' AND 거래처명 LIKE '%${userName}%'`);
      }
      
      console.log(`\n✅ SQL 쿼리 실행 완료: ${rows.length}개 행 조회`);
      if (rows.length > 0) {
        console.log(`   📋 첫 번째 행 샘플:`, {
          정산월: rows[0].settlementMonth || rows[0].정산월,
          사용처: rows[0].merchant || rows[0].사용처,
          사용자: rows[0].userColumn || rows[0].사용자,
          거래처명: rows[0].거래처명 || '',
          금액: rows[0].amount || rows[0].사용금액 || rows[0].출금액,
          비고: (rows[0].note || rows[0].비고 || '').substring(0, 50) + '...'
        });
        
        // 🔥 정산 데이터인 경우 상세 확인
        if (type === 'settled') {
          console.log(`\n   🔍 정산 데이터 상세 분석:`);
          console.log(`   - 전체 조회된 행: ${rows.length}개`);
          
          // 거래처명별 통계
          const 거래처명별통계 = {};
          rows.forEach(row => {
            const 거래처명 = row.거래처명 || row.merchant || '';
            if (거래처명) {
              거래처명별통계[거래처명] = (거래처명별통계[거래처명] || 0) + 1;
            }
          });
          console.log(`   - 거래처명별 통계:`, 거래처명별통계);
          
          // 사용자 필터와 일치하는 행 확인
          if (userName) {
            const 일치하는행 = rows.filter(row => {
              const 거래처명 = row.거래처명 || '';
              return 거래처명 && 거래처명.includes(userName);
            });
            console.log(`   - 거래처명에 "${userName}" 포함된 행: ${일치하는행.length}개`);
            if (일치하는행.length > 0) {
              일치하는행.slice(0, 3).forEach((row, idx) => {
                console.log(`      ${idx + 1}. 정산월: "${row.settlementMonth || row.정산월}", 거래처명: "${row.거래처명 || ''}"`);
              });
            } else {
              console.log(`   ⚠️ 거래처명에 "${userName}"이 포함된 행이 없습니다!`);
              console.log(`   💡 실제 거래처명 샘플:`, Object.keys(거래처명별통계).slice(0, 5));
            }
          }
          
          // 2025-11 데이터 확인
          const rows2025_11 = rows.filter(row => {
            const month = row.settlementMonth || row.정산월 || '';
            return month && String(month).startsWith('2025-11');
          });
          console.log(`   - 2025-11 데이터: ${rows2025_11.length}개`);
          if (rows2025_11.length > 0) {
            rows2025_11.slice(0, 5).forEach((row, idx) => {
              console.log(`      ${idx + 1}. 정산월: "${row.settlementMonth || row.정산월}", 거래처명: "${row.거래처명 || ''}", 사용처: "${row.merchant || row.사용처 || ''}", 금액: ${row.amount || row.출금액 || 0}`);
            });
          } else {
            console.log(`   ⚠️ 2025-11 데이터가 없습니다!`);
            // 정산월별 통계
            const 정산월별통계 = {};
            rows.forEach(row => {
              const month = row.settlementMonth || row.정산월 || '';
              if (month) {
                정산월별통계[month] = (정산월별통계[month] || 0) + 1;
              }
            });
            console.log(`   💡 실제 정산월 분포:`, 정산월별통계);
          }
        }
      } else {
        console.log(`   ⚠️ SQL 쿼리 결과가 비어있습니다.`);
        if (type === 'settled' && userName) {
          console.log(`   💡 사용자 필터("${userName}")에 맞는 데이터가 없을 수 있습니다.`);
          console.log(`   💡 SQL 쿼리 조건: 정산월 >= '2025-11' AND 거래처명 LIKE '%${userName}%'`);
          console.log(`   💡 가능한 원인:`);
          console.log(`      1. SQL 테이블에 2025-11 이후 데이터가 없음`);
          console.log(`      2. 거래처명 컬럼에 "${userName}"이 포함된 데이터가 없음`);
          console.log(`      3. SQL 연결 정보가 잘못됨`);
        }
      }
      
      await pool.close();
      
      // 🔥 정산(SQL) 상세내역 변환 (accountName 포함)
      if (type === 'settled') {
        const detail = [];
        for (const row of rows) {
          // 🔥 paymentDate 형식 변환 (yyyy-mm-dd)
          let paymentDateStr = '';
          let paymentDateObj = null;
          const paymentDate = row.paymentDate || row.반제일 || null;
          if (paymentDate) {
            if (paymentDate instanceof Date) {
              paymentDateObj = paymentDate;
              const year = paymentDate.getFullYear();
              const month = String(paymentDate.getMonth() + 1).padStart(2, '0');
              const day = String(paymentDate.getDate()).padStart(2, '0');
              paymentDateStr = `${year}-${month}-${day}`;
            } else {
              // 문자열인 경우
              const dateStr = String(paymentDate).trim();
              // 이미 yyyy-mm-dd 형식인지 확인
              if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
                paymentDateStr = dateStr; // 이미 올바른 형식이면 그대로 사용
                paymentDateObj = new Date(dateStr);
              } else {
                // 다른 형식이면 Date 객체로 파싱 시도
                const dateObj = new Date(dateStr);
                if (!isNaN(dateObj.getTime())) {
                  paymentDateObj = dateObj;
                  // 로컬 시간을 사용하여 yyyy-mm-dd 형식으로 변환
                  const year = dateObj.getFullYear();
                  const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                  const day = String(dateObj.getDate()).padStart(2, '0');
                  paymentDateStr = `${year}-${month}-${day}`;
                } else {
                  paymentDateStr = dateStr; // 파싱 실패 시 원본 문자열 사용
                }
              }
            }
          }

          // 🔥 계정명 계산 적용 (SQL 정산 데이터 전용 로직)
          // 1번 조건: "월|" 패턴 추출 → match_data_moca 파일의 M열(합계잔액시산표 계정명) 목록과 100% 일치 비교
          const originalNote = row.note || row.비고 || '';
          let accountName = '';
          let 매칭방법 = '매칭실패';
          let 매치율 = 0;
          
          // M열 목록 로드
          const M열합계잔액시산표계정명목록 = await loadMColumnFromMatchDataMoca();
          
          if (M열합계잔액시산표계정명목록 && M열합계잔액시산표계정명목록.length > 0) {
            // "월|" 패턴 추출 후 M열 목록과 비교
            const extractedAccountName = extractAccountNameFromNote(originalNote, M열합계잔액시산표계정명목록);
            if (extractedAccountName) {
              // 1번 조건 성공: 100% 일치
              accountName = extractedAccountName;
              매칭방법 = '월|패턴추출';
              매치율 = 1.0;
            } else {
              // 2번 조건: 100% 일치 없을 경우 "기타" 반환
              accountName = '기타';
              매칭방법 = '매칭실패';
              매치율 = 0;
            }
          } else {
            // M열 목록이 없으면 "기타" 반환
            accountName = '기타';
            매칭방법 = 'M열목록없음';
            매치율 = 0;
          }
          
          // 🔥 2025-11 데이터의 매칭 결과 상세 로그
          const 정산월값 = row.settlementMonth || row.정산월 || '';
          if (정산월값 && 정산월값.startsWith('2025-11')) {
            console.log(`\n🔍 [2025-11 계정명 매칭 결과]`);
            console.log(`   정산월: "${정산월값}"`);
            console.log(`   비고: "${row.note || row.비고 || ''}"`);
            console.log(`   계정명: "${accountName}"`);
            console.log(`   매칭방법: "${매칭방법}"`);
            console.log(`   매치율: ${매치율} (${(매치율 * 100).toFixed(1)}%)`);
          }

          // 🔥 사용처 결정: 2025-11부터의 사용처는 SQL [dbo].[ERP_이체내역조회]의 "사용처" 컬럼에서 가져옴
          // 1. 계정명이 "미지급금_사내"이면 사용처에 거래처명을 넣기
          // 2. 사용처가 null이거나 빈 값이면 거래처명을 사용
          let merchantValue = '';
          
          // 🔥 디버깅: SQL에서 가져온 원본 데이터 확인
          if (정산월값 && 정산월값.startsWith('2025-11')) {
            console.log(`\n🔍 [getSettlementDataFromSQL] 2025-11 데이터 처리 시작:`);
            console.log(`   정산월: "${정산월값}"`);
            console.log(`   계정명: "${accountName}"`);
            console.log(`   SQL 원본 row.merchant: "${row.merchant || '(없음)'}" (타입: ${typeof row.merchant}, null: ${row.merchant === null}, undefined: ${row.merchant === undefined})`);
            console.log(`   SQL 원본 row.사용처: "${row.사용처 || '(없음)'}" (타입: ${typeof row.사용처})`);
            console.log(`   SQL 원본 row.거래처명: "${row.거래처명 || '(없음)'}"`);
            console.log(`   row 객체의 모든 키:`, Object.keys(row).join(', '));
          }
          
          // SQL의 "사용처" 컬럼 값 확인
          const sql사용처값 = row.merchant || row.사용처 || '';
          const 사용처비어있음 = !sql사용처값 || sql사용처값 === null || sql사용처값 === '' || sql사용처값.trim() === '';
          
          if (accountName === '미지급금_사내') {
            // 계정명이 "미지급금_사내"인 경우: 사용처에 거래처명 사용
            merchantValue = row.거래처명 || '';
            if (정산월값 && 정산월값.startsWith('2025-11')) {
              console.log(`   ✅ 계정명이 "미지급금_사내"이므로 거래처명 사용: "${merchantValue}"`);
            }
          } else if (사용처비어있음) {
            // 사용처가 null이거나 빈 값인 경우: 거래처명 사용
            merchantValue = row.거래처명 || '';
            if (정산월값 && 정산월값.startsWith('2025-11')) {
              console.log(`   ✅ 사용처가 비어있으므로 거래처명 사용: "${merchantValue}"`);
            }
          } else {
            // 그 외의 경우: SQL [dbo].[ERP_이체내역조회]의 "사용처" 컬럼 사용
            merchantValue = sql사용처값;
            if (정산월값 && 정산월값.startsWith('2025-11')) {
              console.log(`   ✅ SQL의 "사용처" 컬럼 사용: "${merchantValue}"`);
            }
          }
          
          // 🔥 2025-11 데이터의 merchant 값 확인 로그
          if (정산월값 && 정산월값.startsWith('2025-11')) {
            console.log(`   📋 최종 merchantValue: "${merchantValue || '(없음)'}"`);
          }

          // 🔥 정산월 결정: SQL 정산월 컬럼값 그대로 사용 (지급일 기준 계산 없음)
          const finalSettlementMonth = row.settlementMonth || row.정산월 || null;
          
          const resultItem = {
            month: finalSettlementMonth, // 프론트엔드 필터링을 위해 month 필드 추가
            settlementMonth: finalSettlementMonth,
            paymentDate: paymentDateStr,
            merchant: merchantValue,
            amount: row.amount || row.출금액 || 0,
            note: row.note || row.비고 || '',
            accountName: accountName,
            매칭방법: 매칭방법,  // 매칭방법 정보 추가
            매치율: 매치율,      // 매치율 정보 추가
            isFromSQL: true
          };

          detail.push(resultItem);
        }

        return detail;
      }
      
      // SQL 결과를 표준 형식으로 변환
      // 🔥 미정산의 경우 SQL 변환 루프 안에서 바로 AI 계정매칭으로 accountName 설정
      const sqlDataPromises = rows.map(async (row) => {
        const settlementMonth = row.settlementMonth || row.정산월 || row.month || row.settlement_month || '';
        const normalizedMonth = normalizeSettlementMonth(settlementMonth);
        
        // paymentDate 형식 변환 (미정산: 만기일, 정산: 반제일)
        let paymentDateStr = '';
        const paymentDate = row.paymentDate || row.만기일 || row.반제일 || row.미결발생일 || row.지급일 || row.payment_date;
        if (paymentDate) {
          if (paymentDate instanceof Date) {
            const year = paymentDate.getFullYear();
            const month = String(paymentDate.getMonth() + 1).padStart(2, '0');
            const day = String(paymentDate.getDate()).padStart(2, '0');
            paymentDateStr = `${year}-${month}-${day}`;
          } else {
            // 문자열인 경우 그대로 사용 (이미 yyyy-mm-dd 형식일 것으로 예상)
            paymentDateStr = String(paymentDate).trim();
          }
        }

        // 사용처 설정
        // 🔥 정산 데이터는 거래처명을 사용처로 사용 (SQL 쿼리에서 거래처명으로 필터링했으므로)
        let merchantValue = '';
        if (type === 'settled') {
          merchantValue = row.거래처명 || row.merchant || row.사용처 || '';
        } else {
          merchantValue = row.merchant || row.사용처 || row.거래처명 || '';
        }
        
        // ⭐ SQL 비고에서 계정명 추출
        const originalNote = row.비고 || row.note || "";
        
        let accountNameValue = "";
        
        // 🔥 미정산: 기존 로직 유지
        if (type === 'unsettled') {
          const rawAccountName = extractAccountNameFromSQL(originalNote);
          accountNameValue = rawAccountName;
          
          // 계정명 매칭 적용
          if (!accountNameValue || accountNameValue.trim() === '') {
            try {
              const aiResult = await calculateUnsettledAccountName(originalNote, false, true);
              accountNameValue = aiResult || "";
            } catch (err) {
              accountNameValue = "";
              console.error(`   ⚠️ 계정명 AI 매칭 오류:`, err.message);
          }
        }
        }
        // 🔥 정산(SQL): 계정명 추출은 resultItem 생성 후에 처리 (아래에서 처리)

        const resultItem = {
          month: normalizedMonth || "",
          paymentDate: paymentDateStr,
          merchant: merchantValue,
          amount: Number(row.amount || row.출금액 || row.정산금액 || row["G"] || 0),
          note: originalNote,
          settlementMonth: settlementMonth || normalizedMonth || "",
          isFromSQL: true
        };
        
        // 🔥 계정명 최종 적용
        if (type === 'unsettled') {
          resultItem.accountName = accountNameValue || '-';
        } else if (type === 'settled') {
          // 🔥 정산(SQL) 데이터는 위의 793-890 라인에서 이미 계정명을 계산했으므로 여기서는 처리하지 않음
          // ⚠️ 이 코드는 실행되지 않지만, 혹시 모를 경우를 대비해 주석 처리
          // resultItem.accountName은 이미 위에서 설정되었거나, 793-890 라인의 detail 배열에서 처리됨
          // resultItem.accountName = computedAccountName || "-";
        }
        // 2025-10 이하 정산 데이터(엑셀)는 절대 덮어쓰지 않음 (readExcelAndRespond에서 처리)
        
        return resultItem;
      });

      // 모든 Promise가 완료될 때까지 대기 (기본 데이터 변환 + 미정산/정산 모두 계정명 계산 포함)
      if (type === 'unsettled') {
        console.log(`\n⏳ SQL 데이터 변환 및 미정산 계정명 계산 중... (${rows.length}개 행 처리 중)`);
      } else {
        console.log(`\n⏳ SQL 데이터 변환 및 정산 계정명 계산 중... (${rows.length}개 행 처리 중)`);
      }
      const sqlData = await Promise.all(sqlDataPromises);

      // 🔥 미정산/정산 모두 계정명 계산 완료 확인
      if (type === 'unsettled') {
        console.log(`✅ 미정산 계정명 계산 완료: ${sqlData.length}개 항목`);
      } else {
        console.log(`✅ 정산 계정명 계산 완료: ${sqlData.length}개 항목`);
      }
      
      // 🔥 계산 완료 후 즉시 accountName 확인
      console.log(`\n🔍 Promise.all 완료 후 즉시 accountName 확인 (${type === 'settled' ? '정산' : '미정산'}):`);
      sqlData.slice(0, 5).forEach((item, idx) => {
        const hasAccountName = 'accountName' in item;
        const accountNameValue = item.accountName || '(없음)';
        console.log(`   ${idx + 1}. accountName 필드 존재: ${hasAccountName}, 값: "${accountNameValue}" (타입: ${typeof item.accountName})`);
        console.log(`      정산월: "${item.month || item.settlementMonth || ''}", merchant: "${item.merchant || ''}"`);
        console.log(`      비고: "${(item.note || '').substring(0, 50)}..."`);
        console.log(`      isFromSQL: ${item.isFromSQL || false}`);
      });
      
      // 🔥 정산 데이터인 경우 2025-11 데이터의 accountName 확인
      if (type === 'settled') {
        const settled2025_11 = sqlData.filter(item => {
          const month = item.month || item.settlementMonth || '';
          return month && month.startsWith('2025-11');
        });
        console.log(`\n🔍 2025-11 정산 데이터 accountName 확인:`);
        console.log(`   - 총 ${settled2025_11.length}개 항목`);
        if (settled2025_11.length > 0) {
          settled2025_11.forEach((item, idx) => {
            console.log(`   ${idx + 1}. 정산월: "${item.month || item.settlementMonth}", accountName: "${item.accountName || '(없음)'}", merchant: "${item.merchant || ''}"`);
            console.log(`      비고: "${(item.note || '').substring(0, 50)}..."`);
          });
        } else {
          console.log(`   ⚠️ 2025-11 데이터가 없습니다.`);
        }
      }

      console.log(`\n✅ SQL 데이터 변환 완료: ${sqlData.length}개 항목`);
      if (sqlData.length > 0) {
        console.log(`   📋 변환된 첫 번째 항목 샘플:`, {
          정산월: sqlData[0].settlementMonth || sqlData[0].month,
          사용처: sqlData[0].merchant,
          계정명: sqlData[0].accountName || '(없음)',
          계정명타입: typeof sqlData[0].accountName,
          계정명값: JSON.stringify(sqlData[0].accountName),
          금액: sqlData[0].amount,
          비고: (sqlData[0].note || '').substring(0, 50) + '...'
        });
        
        // 🔥 반환 직전 최종 확인
        console.log(`\n🔍 getSettlementDataFromSQL 반환 직전 최종 확인:`);
        sqlData.slice(0, 3).forEach((item, idx) => {
          const hasAccountName = 'accountName' in item;
          console.log(`   ${idx + 1}. accountName 필드 존재: ${hasAccountName}, 값: "${item.accountName || '(없음)'}" (타입: ${typeof item.accountName})`);
          console.log(`      전체 객체 키: ${Object.keys(item).join(', ')}`);
        });
        
        // 계정명이 없는 항목 확인
        const 계정명없는항목 = sqlData.filter(item => !item.accountName || item.accountName === '' || item.accountName === '-');
        if (계정명없는항목.length > 0) {
          console.log(`   ⚠️ 계정명이 없는 항목: ${계정명없는항목.length}개 / 전체 ${sqlData.length}개`);
          if (계정명없는항목.length <= 5) {
            계정명없는항목.forEach((item, idx) => {
              console.log(`      ${idx + 1}. 비고: "${(item.note || '').substring(0, 50)}...", 계정명: "${item.accountName || '(없음)'}"`);
            });
          }
        } else {
          console.log(`   ✅ 모든 항목에 계정명이 있습니다.`);
        }
      }
      console.log(`${"=".repeat(80)}\n`);
      return sqlData;

    } catch (error) {
      console.error('❌ SQL 데이터 조회 오류:', error);
      console.error('   오류 상세:', error.message);
      if (error.stack) {
        console.error('   스택:', error.stack);
      }
      
      // 연결이 열려있으면 닫기
      if (pool && pool.connected) {
        try {
          await pool.close();
        } catch (closeError) {
          console.error('   연결 종료 중 오류:', closeError.message);
        }
      }
      
      // 오류 발생 시 빈 배열 반환 (엑셀 데이터는 정상적으로 처리되도록)
      return [];
    }
  } catch (error) {
    console.error('❌ SQL 함수 실행 오류:', error);
    return [];
  }
}

function normalizeName(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[()]/g, "")
    .trim();
}

function parseAmountValue(value) {
  if (typeof value === "number" && !isNaN(value)) {
    return value;
  }
  if (typeof value === "string") {
    const numeric = Number(value.replace(/[^0-9.-]/g, ""));
    if (!isNaN(numeric)) {
      return numeric;
    }
  }
  return 0;
}

function formatCurrencyKRW(value) {
  return `${Math.round(value || 0).toLocaleString("ko-KR")}원`;
}

function normalizeMonthString(value) {
  if (value === undefined || value === null) return null;
  const str = String(value).trim();
  if (!str) return null;
  if (/^\d{4}-\d{2}$/.test(str)) {
    return str;
  }
  const digits = str.replace(/[^0-9]/g, "");
  if (digits.length >= 6) {
    const year = digits.slice(0, 4);
    const month = digits.slice(4, 6);
    return `${year}-${month}`;
  }
  return null;
}

function findLatestSettlementMonth(detail = []) {
  let latest = null;
  detail.forEach((item) => {
    const paymentDate = item?.paymentDate || item?.date || null;
    let timestamp = 0;
    let monthLabel = null;

    if (paymentDate) {
      const parsedDate = new Date(paymentDate);
      if (!isNaN(parsedDate.getTime())) {
        timestamp = parsedDate.getTime();
        monthLabel = `${parsedDate.getFullYear()}-${String(parsedDate.getMonth() + 1).padStart(2, "0")}`;
      }
    }

    if (!timestamp) {
      const normalizedMonth = normalizeMonthString(item?.settlementMonth) || normalizeMonthString(item?.month);
      if (normalizedMonth) {
        monthLabel = normalizedMonth;
        const monthDate = Date.parse(`${normalizedMonth}-01T00:00:00Z`);
        timestamp = Number.isNaN(monthDate) ? 0 : monthDate;
      }
    }

    if (monthLabel) {
      if (!latest || timestamp > latest.timestamp) {
        latest = { timestamp, month: monthLabel };
      }
    }
  });
  return latest ? latest.month : null;
}

function findTopSpendingCategory(detail = []) {
  const totals = new Map();
  detail.forEach((item) => {
    const amount = parseAmountValue(item?.amount);
    if (!amount) return;
    const label = item?.accountName || item?.merchant || item?.note || "기타";
    totals.set(label, (totals.get(label) || 0) + amount);
  });

  let result = null;
  totals.forEach((amount, label) => {
    if (!result || amount > result.amount) {
      result = { label, amount };
    }
  });

  return result;
}

// 🔥 거래처명만 확인하는 필터링 함수
function matchUserByMerchant(거래처명값, normalizedUserName) {
  if (!normalizedUserName) return true;
  const target = normalizeName(normalizedUserName);
  if (!target) return true;
  
  if (!거래처명값) return false;
  const candidate = normalizeName(거래처명값);
  return candidate === target || candidate.includes(target);
}

// 기존 함수는 호환성을 위해 유지 (하지만 사용하지 않음)
function matchUserInRow(row, normalizedUserName) {
  if (!normalizedUserName) return true;
  const target = normalizeName(normalizedUserName);
  if (!target) return true;

  return Object.values(row).some(val => {
    if (val === undefined || val === null) return false;
    const candidate = normalizeName(val);
    return candidate === target || candidate.includes(target);
  });
}

// ===================================================
// 📌 각 법인별 데이터를 result 파일로 저장하는 함수
// 기존 파일은 수정하지 않고 새 result 파일만 생성 (병합하지 않음)
// ===================================================
async function saveDataToResultFile(settledDetail, unsettledDetail, sourceFilePath) {
  try {
    // 원본 파일명에서 result 파일명 생성
    const sourceFileName = path.basename(sourceFilePath, path.extname(sourceFilePath));
    const resultFileName = `${sourceFileName}_result.xlsx`;
    const resultPath = getExcelFilePath(`./${resultFileName}`);
    
    console.log(`📝 [${sourceFilePath}] 데이터를 result 파일로 저장 시작`);
    console.log(`   📁 result 파일 경로: ${resultPath}`);
    console.log(`   📁 프로젝트 루트: ${__dirname}`);

    // 🔥 원본 파일의 헤더 구조 확인 (거래처명 컬럼 포함 여부 확인)
    let 원본헤더구조 = null;
    let 거래처명인덱스 = -1;
    try {
      const sourceExcelPath = getExcelFilePath(sourceFilePath);
      if (fs.existsSync(sourceExcelPath)) {
        const 원본워크북 = xlsx.readFile(sourceExcelPath);
        const 원본시트이름 = 원본워크북.SheetNames.find(name => name === "2025" || name.includes("2025")) || 원본워크북.SheetNames[0];
        const 원본시트 = 원본워크북.Sheets[원본시트이름];
        
        if (원본시트) {
          const 원본데이터 = xlsx.utils.sheet_to_json(원본시트, { header: 1, defval: "" });
          
          // 헤더 행 찾기
          let 헤더행 = 0;
          for (let i = 0; i < Math.min(10, 원본데이터.length); i++) {
            const row = 원본데이터[i] || [];
            if (row[0] === "비고" || row[0] === "거래처명" || String(row[0] || "").includes("비고") || String(row[0] || "").includes("거래처명")) {
              헤더행 = i;
              break;
            }
          }
          
          원본헤더구조 = 원본데이터[헤더행] || [];
          거래처명인덱스 = 원본헤더구조.findIndex(h => String(h || "").includes("거래처명"));
          
          console.log(`   📋 원본 파일 헤더 확인: ${원본헤더구조.length}개 컬럼`);
          if (거래처명인덱스 !== -1) {
            console.log(`   ✅ 거래처명 컬럼 발견: 인덱스 ${거래처명인덱스} (${String.fromCharCode(65 + 거래처명인덱스)}열)`);
          } else {
            console.log(`   ⚠️ 원본 파일에 거래처명 컬럼이 없습니다.`);
          }
        }
      }
    } catch (error) {
      console.log(`   ⚠️ 원본 파일 헤더 확인 실패: ${error.message}`);
    }

    // 새 워크북 생성
    const workbook = xlsx.utils.book_new();

    // 1. 정산 시트 생성 (2025 시트)
    if (settledDetail.length > 0) {
      // 헤더 행 생성 (원본 구조 반영)
      const headers = [];
      headers[0] = '비고';    // A열
      headers[6] = '출금액';   // G열
      headers[7] = '지급일';   // H열
      headers[9] = '사용처';   // J열
      headers[10] = '계정명';  // K열
      headers[13] = '정산월';  // N열
      
      // 🔥 원본에 거래처명 컬럼이 있으면 result 파일에도 포함
      if (거래처명인덱스 !== -1) {
        headers[거래처명인덱스] = '거래처명';
        console.log(`   ✅ result 파일에 거래처명 컬럼 추가: 인덱스 ${거래처명인덱스} (${String.fromCharCode(65 + 거래처명인덱스)}열)`);
      }

      // 데이터 행 생성
      const worksheetData = [headers];
      settledDetail.forEach(item => {
        const row = [];
        row[0] = item.note || '';           // A열: 비고
        row[6] = item.amount || 0;          // G열: 출금액
        row[7] = item.paymentDate || '';    // H열: 지급일
        row[9] = item.merchant || '';      // J열: 사용처
        row[10] = item.accountName || '';    // K열: 계정명
        row[13] = item.settlementMonth || item.month || ''; // N열: 정산월
        
        // 🔥 거래처명 컬럼이 있으면 거래처명 값도 포함 (merchant 값 사용)
        if (거래처명인덱스 !== -1) {
          row[거래처명인덱스] = item.merchant || '';
        }
        
        worksheetData.push(row);
      });

      const settledSheet = xlsx.utils.aoa_to_sheet(worksheetData);
      xlsx.utils.book_append_sheet(workbook, settledSheet, "2025");
      console.log(`   ✅ 정산 시트 생성 완료: ${settledDetail.length}개 행`);
    }

    // 2. 미정산 시트 생성 (2025_미정산 시트)
    // 기존 구조: D=거래처명, G=정산 반제할금액, H=만기일, J=비고, AF=계정명, AG=정산월
    if (unsettledDetail.length > 0) {
      // 헤더 행 생성 (기존 구조 유지)
      const headers = [];
      headers[3] = '거래처명';        // D열
      headers[6] = '정산 반제할금액';  // G열
      headers[7] = '만기일';          // H열
      headers[9] = '비고';            // J열
      headers[31] = '계정명';         // AF열 (인덱스 31)
      headers[32] = '정산월';          // AG열 (인덱스 32)

      // 데이터 행 생성
      const worksheetData = [headers];
      // 🔥 미정산 데이터는 비고값으로 계정명을 재계산해야 함 (원본 계정명 무시)
      const 미정산계정명계산Promises = unsettledDetail.map(async (item) => {
        const 비고값 = item.note || '';
        // 비고값으로 계정명 계산 (원본 계정명 무시)
        let 계산된계정명 = '-';
        try {
          const 계산결과 = await calculateUnsettledAccountName(비고값);
          if (계산결과 && 계산결과 !== '-' && 계산결과.trim() !== '') {
            계산된계정명 = 계산결과;
          }
        } catch (error) {
          console.error(`   ⚠️ 미정산 계정명 계산 오류 (비고: "${비고값.substring(0, 50)}..."):`, error.message);
        }
        
        const row = [];
        row[3] = item.merchant || '-';                    // D열: 거래처명
        row[6] = item.amount || 0;                        // G열: 정산 반제할금액
        row[7] = item.paymentDate || '';                  // H열: 만기일
        row[9] = 비고값;                                   // J열: 비고
        row[31] = 계산된계정명;                             // AF열: 계정명 (비고 기준으로 계산된 값, 원본 계정명 무시)
        row[32] = item.settlementMonth || item.month || ''; // AG열: 정산월
        return row;
      });
      
      const 계산된행들 = await Promise.all(미정산계정명계산Promises);
      계산된행들.forEach(row => worksheetData.push(row));

      const unsettledSheet = xlsx.utils.aoa_to_sheet(worksheetData);
      xlsx.utils.book_append_sheet(workbook, unsettledSheet, "2025_미정산");
      console.log(`   ✅ 미정산 시트 생성 완료: ${unsettledDetail.length}개 행`);
    }

    // 3. result 파일 저장
    xlsx.writeFile(workbook, resultPath);
    console.log(`✅ [${sourceFilePath}] result 파일 저장 완료: ${resultPath}`);
    console.log(`   → 기존 파일(${sourceFilePath})은 수정하지 않았습니다.`);
    
  } catch (error) {
    console.error(`❌ [${sourceFilePath}] result 파일 저장 오류: ${error.message}`);
    throw error;
  }
}

// ===================================================
// 📌 정산 상세내역 매치율 정보 추가 파일 생성 함수
// match_data_moca_result2.xlsx 파일 생성 (Q열, R열에 매칭방법, 매치율 추가)
// ===================================================
async function saveSettledMatchRateFile() {
  try {
    const sourceFileName = 'match_data_moca_result.xlsx';
    const resultFileName = 'match_data_moca_result2.xlsx';
    const sourcePath = getExcelFilePath(`./${sourceFileName}`);
    const resultPath = getExcelFilePath(`./${resultFileName}`);
    
    console.log(`\n${"=".repeat(80)}`);
    console.log(`📝 정산 상세내역 매치율 정보 파일 생성 시작`);
    console.log(`   📁 원본 파일: ${sourcePath}`);
    console.log(`   📁 결과 파일: ${resultPath}`);
    
    // 원본 파일 존재 확인
    console.log(`   🔍 원본 파일 존재 확인 중: ${sourcePath}`);
    if (!fs.existsSync(sourcePath)) {
      console.error(`   ❌ 원본 파일이 없습니다: ${sourcePath}`);
      console.error(`   💡 match_data_moca_result.xlsx 파일이 먼저 생성되어야 합니다.`);
      throw new Error(`원본 파일이 없습니다: ${sourcePath}`);
    }
    console.log(`   ✅ 원본 파일 존재 확인: ${sourcePath}`);
    
    // 원본 파일 읽기
    const sourceWorkbook = xlsx.readFile(sourcePath);
    const newWorkbook = xlsx.utils.book_new();
    
    // 모든 시트 처리
    for (const sheetName of sourceWorkbook.SheetNames) {
      const sourceSheet = sourceWorkbook.Sheets[sheetName];
      
      // 시트를 배열로 변환 (모든 데이터 보존)
      const sheetData = xlsx.utils.sheet_to_json(sourceSheet, { header: 1, defval: "" });
      
      // 2025 시트인 경우에만 매치율 추가
      if (sheetName === "2025" || sheetName === "2025moca") {
        console.log(`   📋 ${sheetName} 시트 처리 중... (${sheetData.length}개 행)`);
        
        // 헤더 행 찾기
        let headerRowIndex = 0;
        for (let i = 0; i < Math.min(10, sheetData.length); i++) {
          const row = sheetData[i] || [];
          if (row[0] === "비고" || row[0] === "거래처명" || String(row[0] || "").includes("비고") || String(row[0] || "").includes("거래처명")) {
            headerRowIndex = i;
            break;
          }
        }
        
        // 헤더에 Q열, R열 추가
        const headerRow = sheetData[headerRowIndex] || [];
        if (!headerRow[16]) headerRow[16] = '매칭방법';  // Q열
        if (!headerRow[17]) headerRow[17] = '매치율';    // R열
        
        // 데이터 행 처리 (헤더 다음 행부터)
        const dataPromises = sheetData.slice(headerRowIndex + 1).map(async (row, rowIndex) => {
          // 비고 열 찾기 (A열, 인덱스 0)
          const 비고값 = row[0] || '';
          // 정산월 열 찾기 (N열, 인덱스 13)
          const 정산월값 = row[13] || '';
          
          // 2025-01~2025-10 데이터만 매치율 계산
          const is2025_01_10 = 정산월값 && (
            String(정산월값).startsWith('2025-01') || String(정산월값).startsWith('2025-02') || 
            String(정산월값).startsWith('2025-03') || String(정산월값).startsWith('2025-04') || 
            String(정산월값).startsWith('2025-05') || String(정산월값).startsWith('2025-06') || 
            String(정산월값).startsWith('2025-07') || String(정산월값).startsWith('2025-08') || 
            String(정산월값).startsWith('2025-09') || String(정산월값).startsWith('2025-10')
          );
          
          if (is2025_01_10 && 비고값) {
            try {
              // 비고값으로 계정명 매칭 정보 계산 (상세 정보 포함)
              const aiResult = await calculateUnsettledAccountName(비고값, true);
              row[16] = aiResult.매칭방법 || '없음';  // Q열: 매칭방법
              row[17] = aiResult.매치율 || 0;        // R열: 매치율
            } catch (error) {
              console.error(`   ⚠️ 행 ${headerRowIndex + rowIndex + 2} 매치율 계산 오류:`, error.message);
              row[16] = '오류';  // Q열: 매칭방법
              row[17] = 0;       // R열: 매치율
            }
          } else {
            // 2025-11 이후 데이터는 매치율 정보 없음
            row[16] = row[16] || '';  // Q열: 매칭방법
            row[17] = row[17] || '';  // R열: 매치율
          }
          
          return row;
        });
        
        // 모든 Promise 완료 대기
        const processedData = await Promise.all(dataPromises);
        
        // 헤더 + 처리된 데이터 합치기
        const finalSheetData = [headerRow, ...processedData];
        
        // 새 시트 생성
        const newSheet = xlsx.utils.aoa_to_sheet(finalSheetData);
        xlsx.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
        console.log(`   ✅ ${sheetName} 시트 처리 완료`);
      } else {
        // 다른 시트는 그대로 복사
        const newSheet = xlsx.utils.aoa_to_sheet(sheetData);
        xlsx.utils.book_append_sheet(newWorkbook, newSheet, sheetName);
        console.log(`   ✅ ${sheetName} 시트 복사 완료`);
      }
    }
    
    // 파일 저장
    console.log(`   💾 파일 저장 중...`);
    try {
      xlsx.writeFile(newWorkbook, resultPath);
      console.log(`   ✅ 파일 쓰기 완료: ${resultPath}`);
    } catch (writeError) {
      console.error(`   ❌ 파일 쓰기 오류:`, writeError.message);
      throw writeError;
    }
    
    // 파일이 실제로 생성되었는지 확인
    if (fs.existsSync(resultPath)) {
      const stats = fs.statSync(resultPath);
      console.log(`✅ ${resultFileName} 파일 저장 완료: ${resultPath}`);
      console.log(`   📊 파일 크기: ${stats.size} bytes`);
      console.log(`   📅 파일 생성 시간: ${stats.mtime.toLocaleString('ko-KR')}`);
    } else {
      console.error(`❌ 파일이 생성되지 않았습니다: ${resultPath}`);
      throw new Error(`파일 저장 후 확인 실패: ${resultPath}`);
    }
    console.log("=".repeat(80) + "\n");
    
  } catch (error) {
    console.error(`\n${"=".repeat(80)}`);
    console.error(`❌ 정산 상세내역 매치율 정보 파일 저장 오류`);
    console.error(`   오류 메시지: ${error.message}`);
    if (error.stack) {
      console.error(`   스택: ${error.stack}`);
    }
    console.error("=".repeat(80) + "\n");
    throw error;
  }
}

// ===================================================
// 📌 미정산 상세내역의 AI 반영 계정명 확인 파일 생성 함수

// ===================================================
// 📌 단일 엑셀 파일 처리 함수 (병렬화를 위해 분리)
// ===================================================
async function processSingleExcelFile(excelFilePath, fileIndex, totalFiles, sheetName, normalizedUserName, isUnsettledSheet, period = null) {
  try {
    console.log(`\n${"=".repeat(60)}`);
    console.log(`📁 [${fileIndex + 1}/${totalFiles}] 파일 처리 시작: ${excelFilePath}`);
    console.log(`   📅 조회기간(period): ${period || '없음'} (타입: ${typeof period})`);
    console.log(`${"=".repeat(60)}`);
    
    // 엑셀 파일 경로 처리 (상대 경로를 절대 경로로 변환)
    const excelPath = getExcelFilePath(excelFilePath);
    
    // 파일 존재 확인
    if (!fs.existsSync(excelPath)) {
      console.warn(`⚠️ 파일이 존재하지 않습니다: ${excelPath} (건너뜀)`);
      return { settledDetail: [], monthlyMap: new Map(), unsettledData: [], unsettledAmount: 0, excelFilePath };
    }
    
    console.log(`📋 실제 엑셀 파일 경로: ${excelPath}`);
    console.log(`📋 원본 파일 존재 여부: ${fs.existsSync(excelPath)}`);

    // 🔥 MOCA 파일의 경우 result 파일을 읽기 (2025-01~10 정산월 데이터는 result 파일만 사용)
    let actualExcelPath = excelPath; // 기본값은 원본 파일

    // 🔥 MOCA 파일은 원본 파일에서 직접 읽기 (processExcelFile 호출 안 함)
    if (excelFilePath.includes("match_data_moca")) {
      console.log(`📖 [match_data_moca] 원본 파일에서 직접 읽기 (processExcelFile 건너뜀)`);
      console.log(`   💡 2025-01~2025-10 기간 데이터는 원본 파일의 K열(계정명) 값을 그대로 사용합니다.`);
      // actualExcelPath는 이미 원본 파일 경로로 설정되어 있음
    }
    
    // 시트 목록 가져오기
    let sheetNames = [];
    try {
      sheetNames = await getSheetNames(actualExcelPath);
      console.log(`📋 [${excelFilePath}] 사용 가능한 시트:`, sheetNames);
    } catch (error) {
      console.error(`❌ [${excelFilePath}] 엑셀 파일 읽기 오류: ${error.message} (건너뜀)`);
      return { settledDetail: [], monthlyMap: new Map(), unsettledData: [], unsettledAmount: 0, excelFilePath };
    }
    
    // 지정된 시트에서 데이터 가져오기 (userName으로 필터링)
    let sheetData = [];
    let resultHeaders = [];

    // 🔥 특수 법인(moca 등)은 시트 이름이 다르므로 매핑
    let effectiveSheetName = sheetName;
    if (excelFilePath.includes("match_data_moca")) {
      // moca 법인 시트 매핑
      if (sheetName === "2025" && sheetNames.includes("2025moca")) {
        effectiveSheetName = "2025moca";
      } else if (sheetName === "2025_미정산" && sheetNames.includes("2025_미정산_moca")) {
        effectiveSheetName = "2025_미정산_moca";
      }
    }

    // 🔥 병렬 처리: MOCA 원본 파일 읽기와 실제 파일 읽기를 동시에 실행
    let mocaOriginalData = null;
    let mocaOriginalHeaders = [];
    
    // 병렬로 실행할 작업들 준비
    const readPromises = [];
    
    // MOCA 파일의 경우 원본 파일 읽기 작업 추가
    if (excelFilePath.includes("match_data_moca")) {
      const mocaOriginalPath = path.join(__dirname, 'match_data_moca.xlsx');
      if (fs.existsSync(mocaOriginalPath)) {
        console.log(`📖 [match_data_moca] 원본 파일 읽기 작업 추가 (병렬 처리):`);
        console.log(`   📁 원본 파일 경로: ${mocaOriginalPath}`);
        console.log(`   📄 시트명: 2025moca`);
        readPromises.push(
          getExcelData(
            mocaOriginalPath,
            '2025moca',
            normalizedUserName
          ).then(mocaResult => {
            return { type: 'mocaOriginal', data: mocaResult.data || [], headers: mocaResult.headers || [] };
          }).catch(error => {
            console.error(`❌ [match_data_moca] 원본 파일 읽기 오류: ${error.message}`);
            console.error(`   스택: ${error.stack}`);
            return { type: 'mocaOriginal', data: [], headers: [] };
          })
        );
      } else {
        console.log(`⚠️ [match_data_moca] 원본 파일을 찾을 수 없습니다: ${mocaOriginalPath}`);
      }
    }
    
    // 실제 파일 읽기 작업 추가
    if (sheetNames.includes(effectiveSheetName)) {
      console.log(`📖 [${excelFilePath}] 파일 읽기 작업 추가 (병렬 처리):`);
      console.log(`   📁 읽을 파일 경로: ${actualExcelPath}`);
      console.log(`   📄 시트명: ${effectiveSheetName}`);
      console.log(`   👤 사용자 필터: ${normalizedUserName || '전체'}`);
      readPromises.push(
        getExcelData(
          actualExcelPath,
          effectiveSheetName,
          normalizedUserName
        ).then(result => {
          return { type: 'actual', data: result.data || [], headers: result.headers || [], totalRows: result.totalRows };
        }).catch(error => {
          console.error(`❌ [${excelFilePath}] 시트 데이터 읽기 오류: ${error.message}`);
          console.error(`   스택: ${error.stack}`);
          throw error; // 실제 파일 읽기 실패는 에러로 전파
        })
      );
    } else {
      console.log(`⚠️ [${excelFilePath}] ${effectiveSheetName} 시트를 찾을 수 없습니다. 사용 가능한 시트: ${sheetNames.join(', ')} (건너뜀)`);
      return { settledDetail: [], monthlyMap: new Map(), unsettledData: [], unsettledAmount: 0, excelFilePath };
    }
    
    // 🔥 모든 읽기 작업을 병렬로 실행
    console.log(`🚀 [${excelFilePath}] 엑셀 파일 읽기 병렬 처리 시작: ${readPromises.length}개 작업`);
    const readResults = await Promise.all(readPromises);
    
    // 결과 처리
    for (const result of readResults) {
      if (result.type === 'mocaOriginal') {
        mocaOriginalData = result.data;
        mocaOriginalHeaders = result.headers;
        console.log(`✅ [match_data_moca] 원본 파일에서 ${mocaOriginalData.length}개 행 가져옴 (병렬 처리 완료)`);
      } else if (result.type === 'actual') {
        sheetData = result.data;
        resultHeaders = result.headers;
        console.log(`✅ [${excelFilePath}] ${effectiveSheetName} 시트에서 ${sheetData.length}개 행 가져옴 (전체 행: ${result.totalRows || sheetData.length}개, 병렬 처리 완료)`);
        if (result.totalRows && result.totalRows !== sheetData.length) {
          console.warn(`⚠️ [${excelFilePath}] 경고: 전체 행 수(${result.totalRows})와 반환된 행 수(${sheetData.length})가 다릅니다!`);
        }
        if (sheetData.length > 0) {
          console.log(`   📋 첫 번째 행의 B열(미결발생일) 값: "${sheetData[0]["미결발생일"] || sheetData[0]["Column1"] || '없음'}"`);
          console.log(`📋 [${excelFilePath}] 첫 번째 행 샘플:`, sheetData[0]);
        }
      }
    }

    // 데이터를 개인정산 형식으로 변환 (현재 파일용)
    const settledDetail = [];
    const monthlyMap = new Map(); // 월별 합계 계산용 (현재 파일)

    if (sheetData.length > 0) {
      const firstRow = sheetData[0];
      const headers = Object.keys(firstRow);
      
      // 열 인덱스로 컬럼명 찾기 (기존 로직 유지 - resultHeaders 사용)
      const getColumnNameByIndex = (index) => {
        if (resultHeaders.length > index) {
          const header = resultHeaders[index];
          if (header) {
            return header;
          } else {
            return `Column${index}`;
          }
        } else {
          return `Column${index}`;
        }
      };

      if (isUnsettledSheet) {
        // 🔥 미정산 상세 내역은 엑셀 데이터를 사용하지 않고 SQL 데이터만 사용
        console.log(`⚠️ 미정산 시트는 엑셀에서 읽지 않습니다. SQL 데이터만 사용합니다.`);
        // 미정산 데이터는 빈 배열로 유지 (SQL에서 가져올 예정)
      } else {
        // 정산 시트 매핑 (기존 로직)
        // 정산월 컬럼 찾기 (N열 우선)
      // N열 인덱스 (N열 = 14번째 열, 0-based index: 13)
      const N_COLUMN_INDEX = 13;
      
      let settlementMonthColumnName = null;
      let pendingDateColumnName = null; // 🔥 미결발생일 컬럼 (정산월 없을 때 사용)
      
      // 방법 1: resultHeaders 배열에서 N열 인덱스 확인 (가장 정확)
      if (resultHeaders.length > N_COLUMN_INDEX) {
        const nColumnHeader = resultHeaders[N_COLUMN_INDEX];
        if (nColumnHeader) {
          settlementMonthColumnName = nColumnHeader;
          console.log(`   ✅ N열(인덱스 ${N_COLUMN_INDEX}) 헤더명: "${settlementMonthColumnName}"`);
        } else {
          settlementMonthColumnName = `Column${N_COLUMN_INDEX}`;
          console.log(`   ⚠️ N열(인덱스 ${N_COLUMN_INDEX}) 헤더가 비어있어 Column${N_COLUMN_INDEX}로 설정`);
        }
      } else {
        settlementMonthColumnName = `Column${N_COLUMN_INDEX}`;
        console.log(`   ⚠️ resultHeaders 길이가 ${resultHeaders.length}이므로 Column${N_COLUMN_INDEX}로 설정`);
      }
      
      // 방법 2: 헤더명으로 "정산월" 찾기 (N열이 정산월이 아닌 경우 대비)
      // N열 헤더명이 "정산월"이 아니면 헤더명으로 "정산월" 찾기
      const nColumnHeaderIs정산월 = settlementMonthColumnName && 
        (String(settlementMonthColumnName).trim() === '정산월' || String(settlementMonthColumnName).includes('정산월'));
      
      if (!nColumnHeaderIs정산월) {
        // N열 헤더명이 "정산월"이 아니면 헤더명으로 "정산월" 찾기
        for (const header of headers) {
          if (header && (String(header).trim() === '정산월' || String(header).includes('정산월'))) {
            const foundHeaderIndex = headers.indexOf(header);
            console.log(`   🔍 헤더명 "정산월" 발견: "${header}" (인덱스 ${foundHeaderIndex})`);
            // N열이 아니면 경고
            if (foundHeaderIndex !== N_COLUMN_INDEX) {
              console.warn(`   ⚠️ 경고: "정산월" 헤더가 N열(인덱스 ${N_COLUMN_INDEX})이 아닌 인덱스 ${foundHeaderIndex}에 있습니다.`);
            }
            settlementMonthColumnName = header;
            break;
          }
        }
      }
      
      // 최종 확인: settlementMonthColumnName이 설정되었는지 확인
      if (!settlementMonthColumnName) {
        console.error(`   ❌ 정산월 컬럼명을 찾을 수 없습니다. N열(인덱스 ${N_COLUMN_INDEX})을 사용합니다.`);
        settlementMonthColumnName = `Column${N_COLUMN_INDEX}`;
      }
      
      console.log(`   ✅ 최종 정산월 컬럼명: "${settlementMonthColumnName}"`);
      
      // 방법 3: "미결발생일" 컬럼 찾기 (정산월이 없을 때 yyyy-mm 계산용)
      for (const header of headers) {
        if (header && String(header).includes('미결발생일')) {
          pendingDateColumnName = header;
          break;
        }
      }

      // 열 인덱스 정의 (0-based)
      // A=0, B=1, C=2, D=3, E=4, F=5, G=6, H=7, I=8, J=9, K=10, L=11, M=12, N=13
      const COL_B_INDEX = 1;  // 🔥 미결발생일 (정산월 계산용)
      const COL_G_INDEX = 6;  // 출금액 (정산금액)
      const COL_H_INDEX = 7;  // 지급일
      const COL_J_INDEX = 9;  // 사용처
      const COL_K_INDEX = 10; // 계정명 (거래처명)
      const COL_N_INDEX = 13; // 정산월 (N열, 참고용)

      const 미결발생일ColumnName = getColumnNameByIndex(COL_B_INDEX);
      const 출금액ColumnName = getColumnNameByIndex(COL_G_INDEX);
      const 지급일ColumnName = getColumnNameByIndex(COL_H_INDEX);
      const 사용처ColumnName = getColumnNameByIndex(COL_J_INDEX);
      const 계정명ColumnName = getColumnNameByIndex(COL_K_INDEX);

        // 성능 최적화: 디버깅 로그 제거

      // 🔥 A열 거래처명 확인 (MOCA 파일용, merchant 매핑에 사용)
      const COL_A_INDEX = 0;
      const A열컬럼명 = getColumnNameByIndex(COL_A_INDEX);
      const A열헤더값 = resultHeaders[COL_A_INDEX] || "";
      
      // A열이 거래처명인지 확인
      let 거래처명컬럼키 = null;
      if (A열헤더값 && String(A열헤더값).includes("거래처명")) {
        거래처명컬럼키 = A열컬럼명;
      } else {
        // A열이 거래처명이 아니면 헤더에서 거래처명 찾기
        거래처명컬럼키 = headers.find(h => h && String(h).includes("거래처명")) || 
                        resultHeaders.find(h => h && String(h).includes("거래처명")) ||
                        "거래처명";
      }

      // userName으로 추가 필터링 (정산 시트)
      let filteredSheetData = sheetData;
      const shouldFilterByUserSettled =
        normalizedUserName && sheetData.length > 0;

      if (shouldFilterByUserSettled) {
        const beforeCount = sheetData.length;
        // 🔥 username = 거래처명 기준으로만 필터링

        let 디버그카운트 = 0;
        filteredSheetData = sheetData.filter((row, index) => {
          // 🔥 sheet_to_json으로 변환된 데이터는 헤더명이 키가 됨
          // 따라서 "거래처명" 키로 직접 접근해야 함
          // A열이 거래처명인 경우, 헤더가 "거래처명"이면 키도 "거래처명"
          const 거래처명값 = row["거래처명"] ||  // 1순위: 헤더명으로 직접 접근
                            row[거래처명컬럼키] || 
                            row[A열컬럼명] || 
                            row["Column0"] || 
                            "";
          
          const 매칭결과 = matchUserByMerchant(거래처명값, normalizedUserName);
          
          // 디버깅: 처음 10개 행만 로그
          if (index < 10) {
            console.log(`   [사용자필터] index=${index}, 거래처명="${거래처명값}", 사용자="${normalizedUserName}", 매칭=${매칭결과}`);
          }
          
          return 매칭결과;
        });
        
        console.log(`\n📊 [${excelFilePath}] 사용자 필터링 결과:`);
        console.log(`   👤 사용자: "${normalizedUserName}"`);
        console.log(`   📋 필터링 전: ${beforeCount}개 행`);
        console.log(`   📋 필터링 후: ${filteredSheetData.length}개 행`);
        console.log(`   📅 조회기간: ${period || '없음'}`);
        
        // 🔍 디버깅: 조회기간에 2024가 포함된 경우 정산월별 개수 확인
        if (period && (period.includes('2024') || period.includes('2025-12'))) {
          const 정산월별개수 = {};
          // settlementMonthColumnName이 아직 정의되지 않았을 수 있으므로 N열 인덱스(13)로 직접 접근
          const N_COLUMN_INDEX = 13;
          sheetData.forEach(row => {
            const 정산월값 = row["정산월"] || row[`Column${N_COLUMN_INDEX}`] || '';
            if (정산월값) {
              const 정산월 = String(정산월값).trim();
              if (정산월) {
                if (!정산월별개수[정산월]) {
                  정산월별개수[정산월] = { 전체: 0, 필터링후: 0 };
                }
                정산월별개수[정산월].전체++;
              }
            }
          });
          filteredSheetData.forEach(row => {
            const 정산월값 = row["정산월"] || row[`Column${N_COLUMN_INDEX}`] || '';
            if (정산월값) {
              const 정산월 = String(정산월값).trim();
              if (정산월 && 정산월별개수[정산월]) {
                정산월별개수[정산월].필터링후++;
              }
            }
          });
          console.log(`   📊 [사용자 필터링] 정산월별 개수:`, Object.keys(정산월별개수).sort().map(m => `${m}: 전체=${정산월별개수[m].전체}, 필터링후=${정산월별개수[m].필터링후}`).join(', '));
          
          // 2024-12 데이터가 사용자 필터링에서 제외되었는지 확인
          if (정산월별개수['2024-12']) {
            const 전체2024_12 = 정산월별개수['2024-12'].전체;
            const 필터링후2024_12 = 정산월별개수['2024-12'].필터링후;
            if (전체2024_12 > 0 && 필터링후2024_12 === 0) {
              console.warn(`   ⚠️ [사용자 필터링] 2024-12 데이터가 모두 제외됨: 전체 ${전체2024_12}개 → 필터링 후 0개`);
            } else if (전체2024_12 > 필터링후2024_12) {
              console.warn(`   ⚠️ [사용자 필터링] 2024-12 데이터 일부 제외: 전체 ${전체2024_12}개 → 필터링 후 ${필터링후2024_12}개`);
            }
          }
        }
        
        // 필터링된 데이터의 거래처명 샘플 출력
        if (filteredSheetData.length > 0) {
          const 샘플거래처명 = filteredSheetData.slice(0, 5).map(row => {
            const 거래처명값 = row["거래처명"] || row[거래처명컬럼키] || row[A열컬럼명] || row["Column0"] || "";
            return 거래처명값;
          });
          console.log(`   📋 거래처명 샘플 (처음 5개):`, 샘플거래처명);
        } else if (beforeCount > 0) {
          const 샘플거래처명 = sheetData.slice(0, 5).map(row => {
            const 거래처명값 = row["거래처명"] || row[거래처명컬럼키] || row[A열컬럼명] || row["Column0"] || "";
            return 거래처명값;
          });
          console.warn(`   ⚠️ 필터링 결과가 0개입니다! 원본 거래처명 샘플:`, 샘플거래처명);
        }
      }

      // 각 행 처리
      filteredSheetData.forEach((row, index) => {
        // 🔥 정산월은 N열(정산월) 헤더명 기준으로 읽기 (수기로 입력된 값 사용)
        let settlementMonth = null;
        
        // 정산월 값 가져오기: settlementMonthColumnName을 우선 사용
        let N열정산월값 = null;
        
        // 1순위: settlementMonthColumnName으로 접근 (N열 인덱스 기반 헤더명)
        if (settlementMonthColumnName && row[settlementMonthColumnName] !== undefined) {
          N열정산월값 = row[settlementMonthColumnName];
        }
        // 2순위: "정산월" 헤더명으로 직접 접근 (하위 호환성)
        else if (row["정산월"] !== undefined) {
          N열정산월값 = row["정산월"];
        }
        
        // 디버깅: 처음 5개 행만 로그 출력
        if (index < 5) {
          console.log(`   [${index}] 정산월 읽기: settlementMonthColumnName="${settlementMonthColumnName}", N열값="${N열정산월값}", row["정산월"]="${row["정산월"]}"`);
        }
        
        // 정산월 보정: 숫자 형태도 강제로 텍스트로 처리
        if (N열정산월값 !== null && N열정산월값 !== undefined && N열정산월값 !== "") {
          if (typeof N열정산월값 === "number") {
            N열정산월값 = String(N열정산월값);
          }
          
          // settlementMonth 파싱 실패 방지
          const settlementMonthStr = (N열정산월값 || "").toString().trim();
          const normalizedMonthRaw = settlementMonthStr.replace(/\./g, "-").slice(0, 7);
          if (normalizedMonthRaw && normalizedMonthRaw.length >= 7) {
            settlementMonth = normalizeSettlementMonth(normalizedMonthRaw);
            
            // 디버깅: 처음 5개 행만 로그 출력
            if (index < 5) {
              console.log(`   [${index}] 정산월 파싱: 원본="${N열정산월값}", 정규화전="${normalizedMonthRaw}", 정규화후="${settlementMonth}"`);
            }
          } else {
            // 디버깅: 정산월 파싱 실패
            if (index < 5) {
              console.warn(`   [${index}] ⚠️ 정산월 파싱 실패: 원본="${N열정산월값}", 정규화전="${normalizedMonthRaw}"`);
            }
          }
        } else {
          // 디버깅: 정산월 값이 없음
          if (index < 5) {
            console.warn(`   [${index}] ⚠️ 정산월 값 없음: settlementMonthColumnName="${settlementMonthColumnName}", row keys=${Object.keys(row).join(', ')}`);
          }
        }

        // 출금액 가져오기 (G열) - 정산금액 (헤더명으로 직접 접근)
        let amountValue = null;
        const 출금액Value = row['출금액'] || row[출금액ColumnName];
        if (출금액Value !== undefined && 출금액Value !== null && 출금액Value !== "") {
          amountValue = typeof 출금액Value === 'number' ? 출금액Value : parseFloat(String(출금액Value).replace(/[^0-9.-]/g, ''));
          if (isNaN(amountValue)) {
            amountValue = null;
          }
        }

        // 지급일 가져오기 (H열) - 헤더명으로 직접 접근
        const paymentDate = row["지급일"] || row[지급일ColumnName] || null;

        // 🔥 merchant 값 결정: J열 사용처 컬럼에서 가져오기 (2025-01~2025-10 엑셀 데이터)
        // ⚠️ 엑셀 데이터는 J열(헤더명 "사용처")에서 merchant를 가져옴
        const merchantValue = row['사용처'] || row[사용처ColumnName] || '';

        // 🔥 계정명 가져오기 (K열) - 2025-01~10 엑셀 데이터
        // match_data_moca 파일인 경우 원본 파일의 2025moca 시트 K열에서 그대로 가져옴
        let accountName = '';
        // 정산월값은 settlementMonth만 사용 (normalizedMonth는 아직 정의되지 않음)
        const 정산월값 = settlementMonth || '';
        const is2025_01_10 = 정산월값 && (
          정산월값.startsWith('2025-01') || 정산월값.startsWith('2025-02') || 
          정산월값.startsWith('2025-03') || 정산월값.startsWith('2025-04') || 
          정산월값.startsWith('2025-05') || 정산월값.startsWith('2025-06') || 
          정산월값.startsWith('2025-07') || 정산월값.startsWith('2025-08') || 
          정산월값.startsWith('2025-09') || 정산월값.startsWith('2025-10')
        );
        
        if (excelFilePath.includes("match_data_moca") && is2025_01_10 && mocaOriginalData && mocaOriginalData.length > index) {
          // 2025-01~2025-10 기간: match_data_moca 원본 파일의 2025moca 시트 K열(계정명) 값을 그대로 가져옴
          const originalRow = mocaOriginalData[index];
          if (originalRow) {
            const COL_K_INDEX = 10;
            const 계정명ColumnName = mocaOriginalHeaders[COL_K_INDEX] || `Column${COL_K_INDEX}`;
            accountName = originalRow[계정명ColumnName] || originalRow["Column10"] || originalRow["계정명"] || '';
          }
        } else {
          // 기존 로직 (다른 파일의 경우)
          if (resultHeaders[10] && String(resultHeaders[10]).trim() === '계정명') {
            accountName = row[resultHeaders[10]] || row["Column10"] || '';
          } else {
            accountName = row["Column10"] || '';
          }
        }
        
        // 성능 최적화: 디버깅 로그 제거

        // 비고 컬럼 찾기 (I열 또는 헤더명으로)
        const COL_I_INDEX = 8;
        const 비고ColumnName = getColumnNameByIndex(COL_I_INDEX);
        const originalNote = row["비고"] || row[비고ColumnName] || row["적요"] || row["내용"] || "";

        // 지급일 형식 변환 (H열) - MOCA 원본 파일의 H열 값을 yyyy-mm-dd 형식으로 변환 (타임존 문제 방지)
        let paymentDateStr = null;
        if (paymentDate) {
          if (typeof paymentDate === "number") {
            // Excel 날짜 형식 변환 (숫자 → yyyy-mm-dd)
            // Excel 기준일: 1899-12-30
            const excelEpoch = new Date(1899, 11, 30);
            const jsDate = new Date(excelEpoch.getTime() + paymentDate * 24 * 60 * 60 * 1000);
            // 로컬 시간을 사용하여 yyyy-mm-dd 형식으로 변환 (toISOString()은 UTC로 변환되어 날짜가 변경될 수 있음)
            const year = jsDate.getFullYear();
            const month = String(jsDate.getMonth() + 1).padStart(2, '0');
            const day = String(jsDate.getDate()).padStart(2, '0');
            paymentDateStr = `${year}-${month}-${day}`;
          } else {
            // 문자열인 경우
            const dateStr = String(paymentDate).trim();
            // 이미 yyyy-mm-dd 형식인지 확인
            if (/^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
              paymentDateStr = dateStr; // 이미 올바른 형식이면 그대로 사용
            } else {
              // 다른 형식이면 Date 객체로 파싱 시도
              const dateObj = new Date(dateStr);
              if (!isNaN(dateObj.getTime())) {
                // 로컬 시간을 사용하여 yyyy-mm-dd 형식으로 변환
                const year = dateObj.getFullYear();
                const month = String(dateObj.getMonth() + 1).padStart(2, '0');
                const day = String(dateObj.getDate()).padStart(2, '0');
                paymentDateStr = `${year}-${month}-${day}`;
              } else {
                paymentDateStr = dateStr; // 파싱 실패 시 원본 문자열 사용
              }
            }
          }
        }

        const normalizedMonth = normalizeSettlementMonth(settlementMonth);

        // 정산월 설정
        let month = normalizedMonth;

        // 금액이 있으면 데이터 추가
        if (amountValue !== null && !isNaN(amountValue)) {
          // 🔥 정산월이 없으면 데이터 추가하지 않음 (오류 방지)
          if (!month) {
            console.warn(`   ⚠️ [${index}] 정산월이 없어 데이터 추가 건너뜀: 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
            return; // 정산월이 없으면 건너뛰기
          }

          // 🔥 정산월 필터링: 조회기간에 맞는 데이터만 처리 (엑셀은 ~2025-10까지만)
          // 조회기간은 필수값이므로 항상 있음
          const item정산월 = normalizedMonth || settlementMonth || '';
          
          // 🔥 2025-11 이상 데이터는 엑셀에서 제외 (SQL에서 가져옴)
          if (item정산월 && item정산월 >= '2025-11') {
            if (index < 10) {
              console.log(`   [${index}] [엑셀 필터] 2025-11 이상 데이터 제외 (SQL에서 가져옴): 정산월="${item정산월}", 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
            }
            return; // 2025-11 이상은 엑셀에서 제외
          }
          
          // period 파싱: "2024-01 ~ 2025-02" 또는 "2024-01 - 2025-02" 형식
          // 🔥 period가 없으면 필터링하지 않음 (모든 데이터 통과)
          if (!period) {
            // period가 없으면 필터링하지 않음
            if (index < 10) {
              console.log(`   [${index}] [정산월 필터] period가 없어 필터링 건너뜀: 정산월="${item정산월}", 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
            }
            // period가 없으면 필터링하지 않고 통과
          } else {
            // 🔥 period 형식 다양하게 지원: "2024-01 ~ 2025-12", "2024-01 - 2025-12", "2024-01~2025-12"
            const periodMatch = period.match(/(\d{4})-(\d{2})\s*[-~]\s*(\d{4})-(\d{2})/);
            if (periodMatch && item정산월) {
              const [, startYear, startMonth, endYear, endMonth] = periodMatch;
              const startMonthKey = `${startYear}-${startMonth}`;
              const endMonthKey = `${endYear}-${endMonth}`;
              
              // 정산월이 조회기간 범위에 있는지 확인 (문자열 비교로 정확하게)
              const isInRange = item정산월 >= startMonthKey && item정산월 <= endMonthKey;
              
              // 🔍 디버깅: 2024-12 데이터 확인 (조회기간 2024-01~2025-12인 경우)
              if (item정산월 === '2024-12' && startMonthKey <= '2024-12' && endMonthKey >= '2024-12') {
                if (index < 20) {
                  console.log(`   🔍 [${index}] [정산월 필터] 2024-12 데이터 확인: 정산월="${item정산월}", 조회기간=${startMonthKey}~${endMonthKey}, 포함=${isInRange}, 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
                }
              }
              
              if (!isInRange) {
                // 조회기간 범위 밖의 데이터는 건너뛰기
                if (index < 10) {
                  console.log(`   [${index}] [정산월 필터] 조회기간 범위 밖 데이터 건너뜀: 정산월="${item정산월}", 조회기간=${startMonthKey}~${endMonthKey}, 포함=${isInRange}, 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
                }
                return; // 조회기간 범위 밖이면 건너뛰기
              } else {
                // 디버깅: 조회기간 2024-01~2025-12인 경우 2024-12 데이터 확인
                if (item정산월 === '2024-12' && startMonthKey === '2024-01' && endMonthKey === '2025-12') {
                  if (index < 20) {
                    console.log(`   ✅ [${index}] [정산월 필터] 2024-12 데이터 포함: 정산월="${item정산월}", 조회기간=${startMonthKey}~${endMonthKey}, 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
                  }
                }
              }
            } else if (!periodMatch) {
              // period 형식이 잘못된 경우
              console.warn(`   ⚠️ [${index}] [정산월 필터] period 형식 오류: "${period}", 정산월="${item정산월}"`);
            } else if (!item정산월) {
              // 정산월이 없는 경우
              if (index < 10) {
                console.warn(`   ⚠️ [${index}] [정산월 필터] 정산월이 없음: period="${period}", 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
              }
            }
          }

          // 디버깅: match_data_moca 파일의 경우 더 자세한 로그
          if (excelFilePath.includes("match_data_moca")) {
            if (index < 10) {
              console.log(`   [${index}] [match_data_moca] 데이터 추가: 정산월="${month}", 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
            }
          } else if (index < 5) {
            console.log(`   [${index}] 데이터 추가: 정산월="${month}", 출금액=${amountValue}, 거래처명="${row["거래처명"] || ''}"`);
          }

          settledDetail.push({
            month: month, // normalizedMonth 대신 month 사용 (확실한 값)
            paymentDate: paymentDateStr,
            merchant: merchantValue,
            amount: Number(row['출금액']) || 0, // G열 출금액 강제 적용
            note: originalNote,
            settlementMonth: settlementMonth || month, // 원본 정산월 값도 저장
            accountName: accountName || '',  // K열 계정명 (1598-1603번 줄에서 계산된 값 사용)
            isFromSQL: false
          });

          // 월별 합계 계산 (미정산 데이터 제외)
          // month는 이미 normalizedMonth로 설정되어 있고, null 체크는 위에서 했으므로 여기서는 항상 값이 있음
          // 미정산 데이터 제외 (정산월에 "미정산"이 포함된 경우 제외)
          if (month.includes('미정산') || month.includes('_미정산')) {
            return; // 이 행은 건너뛰기
          }
          
          if (monthlyMap.has(month)) {
            monthlyMap.set(month, monthlyMap.get(month) + amountValue);
          } else {
            monthlyMap.set(month, amountValue);
          }
          
          // 디버깅: 처음 5개 행만 로그 출력
          if (index < 5) {
            console.log(`   [${index}] 월별 합계 업데이트: 정산월="${month}", 금액=${amountValue}, 누적합계=${monthlyMap.get(month)}`);
          }
        }
      });
      }
    }

    // 🔥 현재 파일의 정산 데이터 합계 계산
    const currentFileTotal = settledDetail.reduce((sum, item) => {
        const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
        return sum + amount;
      }, 0);
      
    // 🔥 match_data_moca 파일의 경우 2025-01~2025-10 데이터 통계 출력
    if (excelFilePath.includes("match_data_moca")) {
      const moca2025_01_10 = settledDetail.filter(item => {
        const month = item.month || item.settlementMonth || '';
        return month && (
          month.startsWith('2025-01') || month.startsWith('2025-02') || 
          month.startsWith('2025-03') || month.startsWith('2025-04') || 
          month.startsWith('2025-05') || month.startsWith('2025-06') || 
          month.startsWith('2025-07') || month.startsWith('2025-08') || 
          month.startsWith('2025-09') || month.startsWith('2025-10')
        );
      });
      
      const moca2025_01_10Total = moca2025_01_10.reduce((sum, item) => {
        const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
        return sum + amount;
      }, 0);
      
      console.log(`\n${"=".repeat(80)}`);
      console.log(`📊 [match_data_moca] 2025-01~2025-10 데이터 통계:`);
      console.log(`   👤 사용자 필터: "${normalizedUserName || '전체'}"`);
      console.log(`   ✅ 2025-01~2025-10 데이터: ${moca2025_01_10.length}개 항목, 합계: ${moca2025_01_10Total.toLocaleString()}원`);
      console.log(`   📋 전체 데이터: ${settledDetail.length}개 항목, 합계: ${currentFileTotal.toLocaleString()}원`);
      
      // 월별 통계 출력
      const mocaMonthlyMap = new Map();
      moca2025_01_10.forEach(item => {
        const month = item.month || item.settlementMonth || '';
        if (month) {
          const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
          if (mocaMonthlyMap.has(month)) {
            mocaMonthlyMap.set(month, mocaMonthlyMap.get(month) + amount);
          } else {
            mocaMonthlyMap.set(month, amount);
          }
        }
      });
      
      console.log(`   📅 월별 통계:`);
      Array.from(mocaMonthlyMap.entries())
        .sort((a, b) => a[0].localeCompare(b[0]))
        .forEach(([month, amount]) => {
          console.log(`      ${month}: ${amount.toLocaleString()}원 (${moca2025_01_10.filter(item => (item.month || item.settlementMonth) === month).length}개 항목)`);
        });
      
      // 2025-01 상세 확인
      const moca2025_01 = moca2025_01_10.filter(item => {
        const month = item.month || item.settlementMonth || '';
        return month && month.startsWith('2025-01');
      });
      
      if (moca2025_01.length > 0) {
        console.log(`\n   🔍 2025-01 상세 내역 (${moca2025_01.length}개 항목):`);
        moca2025_01.forEach((item, idx) => {
          console.log(`      ${idx + 1}. 정산월="${item.month}", 사용처="${item.merchant}", 금액=${item.amount.toLocaleString()}원`);
        });
        const moca2025_01Total = moca2025_01.reduce((sum, item) => sum + (item.amount || 0), 0);
        console.log(`      💰 2025-01 합계: ${moca2025_01Total.toLocaleString()}원`);
      }
      
      console.log(`${"=".repeat(80)}\n`);
    }
      
    console.log(`✅ [${excelFilePath}] ${settledDetail.length}개 항목 변환 완료 (합계: ${currentFileTotal.toLocaleString()}원)`);

    // 🔥 미정산 시트명 설정
    const UNSETTLED_SHEET_NAME = "2025_미정산";

    // 🔥 미정산 데이터 읽기 (현재 파일)
    let unsettledData = [];
    let unsettledAmount = 0;

    if (isUnsettledSheet) {
      // 🔥 미정산 시트는 엑셀에서 읽지 않음 (SQL 데이터만 사용)
      console.log(`⚠️ [${excelFilePath}] 미정산 시트는 엑셀에서 읽지 않습니다. SQL 데이터만 사용합니다.`);
      unsettledData = []; // 빈 배열로 유지
      unsettledAmount = 0;
    } else {
      // 🔥 정산 시트인 경우: 미정산 시트는 엑셀에서 읽지 않음 (SQL 데이터만 사용)
      console.log(`⚠️ [${excelFilePath}] 미정산 시트는 엑셀에서 읽지 않습니다. SQL 데이터만 사용합니다.`);
      // unsettledData는 빈 배열로 유지 (SQL에서 가져올 예정)
    }

    console.log(`✅ [${excelFilePath}] 정산 데이터 로드 완료`);
    console.log(`📊 [${excelFilePath}] 월별 요약: ${monthlyMap.size}개, 상세 내역: ${settledDetail.length}개`);

    // 🔥 각 파일별로 별도 result 파일 생성 (병합하지 않음, 비동기 처리)
    // OpenAI로 처리된 파일(match_data_moca)은 이미 result 파일이 있으므로 skip
    // 파일 생성을 비동기로 실행하여 병렬 처리 속도 향상
    if (
      !isUnsettledSheet &&
      settledDetail.length > 0 &&
      !excelFilePath.includes("match_data_moca") &&
      !SKIP_FILE_WRITE
    ) {
      // 파일 생성을 비동기로 실행 (응답 대기 시간 단축)
      setImmediate(async () => {
        try {
          console.log(`📝 [${excelFilePath}] [비동기] result 파일 생성 시작...`);
          await saveDataToResultFile(settledDetail, unsettledData, excelFilePath);
          console.log(`✅ [${excelFilePath}] [비동기] result 파일 생성 완료`);
        } catch (error) {
          console.error(`❌ [${excelFilePath}] [비동기] result 파일 저장 중 오류 발생:`);
          console.error(`   오류 내용: ${error.message}`);
          console.error(`   스택: ${error.stack}`);
        }
      });
      console.log(`📝 [${excelFilePath}] result 파일 생성 예약됨 (비동기 처리)`);
    } else if (excelFilePath.includes("match_data_moca") && !SKIP_FILE_WRITE) {
      console.log(`📋 [${excelFilePath}] OpenAI result 파일이 이미 생성되었으므로 skip`);
      
      // 🔥 MOCA 파일의 경우 매치율 정보 파일 생성 (match_data_moca_result2.xlsx, 비동기 처리)
      setImmediate(async () => {
        try {
          console.log(`\n${"=".repeat(80)}`);
          console.log(`📝 [${excelFilePath}] [비동기] 매치율 정보 파일 생성 시작...`);
          console.log(`   원본 파일 확인: match_data_moca_result.xlsx`);
          await saveSettledMatchRateFile();
          console.log(`✅ [${excelFilePath}] [비동기] 매치율 정보 파일 생성 완료`);
          console.log(`${"=".repeat(80)}\n`);
        } catch (error) {
          console.error(`\n${"=".repeat(80)}`);
          console.error(`❌ [${excelFilePath}] [비동기] 매치율 정보 파일 생성 중 오류 발생:`);
          console.error(`   오류 내용: ${error.message}`);
          console.error(`   스택: ${error.stack}`);
          console.error(`${"=".repeat(80)}\n`);
        }
      });
      console.log(`📝 [${excelFilePath}] 매치율 정보 파일 생성 예약됨 (비동기 처리)`);
    } else {
      console.log(`📋 [${excelFilePath}] result 파일 생성 조건 불만족: isUnsettledSheet=${isUnsettledSheet}, settledDetail.length=${settledDetail.length}`);
    }

    return { settledDetail, monthlyMap, unsettledData, unsettledAmount, excelFilePath };
  } catch (error) {
    console.error(`❌ [${excelFilePath}] 파일 처리 중 오류 발생:`, error.message);
    console.error(`   스택: ${error.stack}`);
    return { settledDetail: [], monthlyMap: new Map(), unsettledData: [], unsettledAmount: 0, excelFilePath };
  }
}

// ===================================================
// 📌 readExcelAndRespond 함수
// 엑셀 파일에서 데이터를 읽어서 응답하는 공통 함수
// 🔥 여러 법인의 엑셀 파일을 병렬로 병합하여 처리
// ===================================================
async function readExcelAndRespond(res, sheetName, userName, period = null) {
  try {
    const normalizedUserName = userName ? String(userName).trim() : null;
    console.log("📌 readExcelAndRespond 호출됨");
    console.log(`📋 시트명: ${sheetName}, userName: ${normalizedUserName || '전체'}, period: ${period || '없음'}`);

    // 🔥 조회 시 과거 캐시 삭제
    responseCache.clear();
    console.log(`🗑️ 조회 시 과거 캐시 삭제 완료`);

    // 캐시 키 생성 (응답 저장용)
    const cacheKey = `${sheetName}__${normalizedUserName || 'ALL'}__${period || 'ALL'}`;

    // 🔥 처리할 엑셀 파일 목록 생성 (MOCA 파일만 처리)
    console.log("📋 ENV ADDITIONAL_EXCEL_FILES:", process.env.ADDITIONAL_EXCEL_FILES);
    // MOCA 파일만 처리 (기존법인 파일은 제외)
    const excelFilePaths = ADDITIONAL_EXCEL_FILES.filter(file => file.toLowerCase().includes("moca"));
    console.log(`📋 처리할 엑셀 파일 목록 (MOCA만, ${excelFilePaths.length}개):`, excelFilePaths);

    // 모든 파일의 데이터를 병합할 변수
    const allSettledDetail = [];
    const allMonthlyMap = new Map(); // 월별 합계 계산용 (병합)
    const allUnsettledData = [];
    let allUnsettledAmount = 0;

    // 시트명에 따라 다른 매핑 사용 (루프 밖에서 정의)
    const isUnsettledSheet = sheetName === "2025_미정산";

    // 🔥 병렬 처리: 엑셀 파일 읽기와 SQL 쿼리를 동시에 실행
    console.log(`\n🚀 병렬 처리 시작: ${excelFilePaths.length}개 엑셀 파일 + SQL 쿼리 동시 처리`);
    
    // 엑셀 파일 처리 Promise
    const fileProcessingPromises = excelFilePaths.map((excelFilePath, fileIndex) => 
      processSingleExcelFile(excelFilePath, fileIndex, excelFilePaths.length, sheetName, normalizedUserName, isUnsettledSheet, period)
    );
    
    // SQL 쿼리 Promise들 (정산 + 미정산)
    // 🔥 조회기간에 포함된 정산월 중 2025-11 이상이 있으면 SQL에서 데이터 가져오기
    // 예: 조회기간 2025-10~2025-11 → 정산월 2025-10은 엑셀, 2025-11은 SQL
    let sqlSettledPromise = Promise.resolve([]);
    let sqlUnsettledPromise = Promise.resolve([]);
    
    if (period) {
      const periodMatch = period.match(/(\d{4})-(\d{2})\s*~\s*(\d{4})-(\d{2})/);
      if (periodMatch) {
        const [, startYear, startMonth, endYear, endMonth] = periodMatch;
        const startMonthKey = `${startYear}-${startMonth}`;
        const endMonthKey = `${endYear}-${endMonth}`;
        
        // 조회기간에 포함된 정산월 중 2025-11 이상이 있는지 확인
        // 종료 월이 2025-11 이상이면 SQL에서 데이터 가져오기
        if (endMonthKey >= '2025-11') {
          console.log(`📊 SQL 쿼리 병렬 실행 준비: 정산 데이터 + 미정산 데이터 (조회기간: ${startMonthKey}~${endMonthKey}, 정산월 2025-11 이상 포함)`);
          sqlSettledPromise = getSettlementDataFromSQL(normalizedUserName, 'settled', period);
          sqlUnsettledPromise = getSettlementDataFromSQL(normalizedUserName, 'unsettled');
        } else {
          console.log(`📊 SQL 쿼리 건너뜀: 조회기간(${startMonthKey}~${endMonthKey})에 정산월 2025-11 이상이 없으므로 SQL에서 데이터를 가져오지 않습니다.`);
        }
      }
    }
    
    // 🔥 모든 작업을 병렬로 실행 (엑셀 파일 읽기 + SQL 정산 쿼리 + SQL 미정산 쿼리)
    const [fileResults, sqlSettledDetail, sqlUnsettledDetail] = await Promise.allSettled([
      Promise.allSettled(fileProcessingPromises),
      sqlSettledPromise,
      sqlUnsettledPromise
    ]);
    
    // 엑셀 파일 결과 병합
    const fileResultsArray = fileResults.status === 'fulfilled' ? fileResults.value : [];
    fileResultsArray.forEach((result, index) => {
      if (result.status === 'fulfilled') {
        const { settledDetail, monthlyMap, unsettledData, unsettledAmount } = result.value;
        allSettledDetail.push(...settledDetail);
        
        // 월별 합계 병합 (미정산 데이터 제외)
        monthlyMap.forEach((amount, month) => {
          // 미정산 데이터 제외 (정산월에 "미정산"이 포함된 경우 제외)
          if (month && (month.includes('미정산') || month.includes('_미정산'))) {
            return;
          }
          if (allMonthlyMap.has(month)) {
            allMonthlyMap.set(month, allMonthlyMap.get(month) + amount);
          } else {
            allMonthlyMap.set(month, amount);
          }
        });
        
        allUnsettledData.push(...unsettledData);
        allUnsettledAmount += unsettledAmount;
      } else {
        console.error(`❌ 파일 처리 실패 [${excelFilePaths[index]}]:`, result.reason);
      }
    });

    console.log(`\n✅ 모든 파일 병렬 처리 완료!`);
    console.log(`📊 전체 정산 상세 내역(필터 전): ${allSettledDetail.length}개`);
    console.log(`📊 전체 미정산 상세 내역(필터 전): ${allUnsettledData.length}개, 합계: ${allUnsettledAmount}`);
    
    // 🔍 디버깅: 조회기간에 포함된 정산월별 데이터 확인
    if (period) {
        const periodMatch = period.match(/(\d{4})-(\d{2})\s*~\s*(\d{4})-(\d{2})/);
        if (periodMatch) {
            const [, startYear, startMonth, endYear, endMonth] = periodMatch;
            const startMonthKey = `${startYear}-${startMonth}`;
            const endMonthKey = `${endYear}-${endMonth}`;
            
            // 정산월별 데이터 확인
            const byMonth = {};
            allSettledDetail.forEach(item => {
                const month = item.month || item.settlementMonth || '없음';
                if (!byMonth[month]) {
                    byMonth[month] = [];
                }
                byMonth[month].push(item);
            });
            console.log(`📊 [서버 원본 데이터] 정산월별 개수:`, Object.keys(byMonth).sort().map(m => `${m}: ${byMonth[m].length}개`).join(', '));
            
            // 조회기간에 포함된 정산월별로 확인
            const periodMonths = [];
            let currentYear = parseInt(startYear);
            let currentMonth = parseInt(startMonth);
            const endYearInt = parseInt(endYear);
            const endMonthInt = parseInt(endMonth);
            
            while (currentYear < endYearInt || (currentYear === endYearInt && currentMonth <= endMonthInt)) {
                const monthKey = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;
                periodMonths.push(monthKey);
                currentMonth++;
                if (currentMonth > 12) {
                    currentMonth = 1;
                    currentYear++;
                }
            }
            
            console.log(`📊 [조회기간 포함 정산월]: ${periodMonths.join(', ')}`);
            periodMonths.forEach(monthKey => {
                const count = byMonth[monthKey] ? byMonth[monthKey].length : 0;
                const source = byMonth[monthKey] && byMonth[monthKey].length > 0 
                    ? (byMonth[monthKey][0].isFromSQL ? 'SQL' : '엑셀')
                    : '없음';
                console.log(`   ${monthKey}: ${count}개 (${source})`);
            });
        }
    }
    
    // 🔍 디버깅: 조회기간 2025-01~2025-02인 경우 상세 확인 (기존 로직 유지)
    if (period && period.includes('2025-01') && period.includes('2025-02')) {
        // 정산월별 데이터 확인
        const byMonth = {};
        allSettledDetail.forEach(item => {
            const month = item.month || item.settlementMonth || '없음';
            if (!byMonth[month]) {
                byMonth[month] = [];
            }
            byMonth[month].push(item);
        });
        console.log(`📊 [서버 원본 데이터] 정산월별 개수:`, Object.keys(byMonth).sort().map(m => `${m}: ${byMonth[m].length}개`).join(', '));
        
        // 지급일별 데이터 확인
        const byPaymentDate = {};
        allSettledDetail.forEach(item => {
            if (item.paymentDate) {
                const paymentDateStr = String(item.paymentDate).trim();
                let paymentYearMonth = '';
                if (/^\d{4}-\d{2}-\d{2}$/.test(paymentDateStr)) {
                    paymentYearMonth = paymentDateStr.substring(0, 7);
                } else if (/^\d{4}-\d{2}$/.test(paymentDateStr)) {
                    paymentYearMonth = paymentDateStr;
                }
                if (paymentYearMonth) {
                    if (!byPaymentDate[paymentYearMonth]) {
                        byPaymentDate[paymentYearMonth] = [];
                    }
                    byPaymentDate[paymentYearMonth].push(item);
                }
            }
        });
        console.log(`📊 [서버 원본 데이터] 지급일(YYYY-MM)별 개수:`, Object.keys(byPaymentDate).sort().map(d => `${d}: ${byPaymentDate[d].length}개`).join(', '));
        
        // 2024-12 지급일 데이터 확인
        const payment2024_12 = allSettledDetail.filter(item => {
            if (!item.paymentDate) return false;
            const paymentDateStr = String(item.paymentDate).trim();
            let paymentYearMonth = '';
            if (/^\d{4}-\d{2}-\d{2}$/.test(paymentDateStr)) {
                paymentYearMonth = paymentDateStr.substring(0, 7);
            } else if (/^\d{4}-\d{2}$/.test(paymentDateStr)) {
                paymentYearMonth = paymentDateStr;
            }
            return paymentYearMonth === '2024-12';
        });
        console.log(`📊 [서버 원본 데이터] 지급일 2024-12인 데이터: ${payment2024_12.length}개`);
        if (payment2024_12.length > 0) {
            console.log(`   📋 샘플 (처음 5개):`, payment2024_12.slice(0, 5).map(item => ({
                정산월: item.month || item.settlementMonth,
                지급일: item.paymentDate,
                사용처: item.merchant,
                금액: item.amount,
                출처: item.isFromSQL ? 'SQL' : '엑셀'
            })));
        }
        
        // 2024-12 정산월 데이터 확인
        const month2024_12 = allSettledDetail.filter(item => {
            const month = item.month || item.settlementMonth || '';
            return month === '2024-12' || month.startsWith('2024-12');
        });
        console.log(`📊 [서버 원본 데이터] 정산월 2024-12인 데이터: ${month2024_12.length}개`);
        if (month2024_12.length > 0) {
            console.log(`   📋 샘플 (처음 5개):`, month2024_12.slice(0, 5).map(item => ({
                정산월: item.month || item.settlementMonth,
                지급일: item.paymentDate,
                사용처: item.merchant,
                금액: item.amount,
                출처: item.isFromSQL ? 'SQL' : '엑셀'
            })));
        }
    }

      // 🔥 필터링 전 전체 정산 데이터 합계 계산
      const beforeFilterTotal = Array.isArray(allSettledDetail) ? allSettledDetail.reduce((sum, item) => {
        try {
          const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
          return sum + amount;
        } catch (e) {
          console.error('⚠️ 합계 계산 중 오류:', e, item);
          return sum;
        }
      }, 0) : 0;
      console.log(`📊 필터링 전 전체 정산 합계: ${beforeFilterTotal.toLocaleString()}원 (${Array.isArray(allSettledDetail) ? allSettledDetail.length : 0}개 항목)`);

      // 🔥 SQL에서 지급일 기준으로 데이터 가져오기 (병렬 처리 완료)
      // 정산 상세내역: [dbo].[ERP_이체내역조회]
      console.log(`\n📊 SQL 정산 데이터 조회 완료 (병렬 처리)`);
      const sqlSettledData = sqlSettledDetail.status === 'fulfilled' ? sqlSettledDetail.value : [];
      
      if (sqlSettledDetail.status === 'rejected') {
        console.error(`❌ SQL 정산 데이터 조회 실패:`, sqlSettledDetail.reason);
      }
      
      if (sqlSettledData.length > 0) {
        console.log(`✅ SQL에서 정산 ${sqlSettledData.length}개 데이터 조회 완료`);
        
        // 🔥 SQL 데이터 상세 확인
        // 성능 최적화: 상세 로그 제거
        
        // 🔵 SQL 정산 데이터 처리
        // getSettlementDataFromSQL에서 이미 모든 필드(merchant, accountName 등)가 제대로 설정되어 반환되므로
        // 그대로 사용하면 됨 (중복 처리 불필요)
        // 🔥 단, 2025-01~2025-10 기간 데이터는 엑셀에서 이미 가져왔으므로 제외 (중복 방지)
        const beforeCount = allSettledDetail.length;
        let sql2025_01_10ExcludedCount = 0;
        let sql2025_11_plusCount = 0;
        
        for (const row of sqlSettledData) {
          const row정산월 = row.settlementMonth || row.정산월 || '';
          
          // 🔥 2025-01~2025-10 기간 데이터는 엑셀에서 이미 가져왔으므로 제외
          if (row정산월 && (
            row정산월.startsWith('2025-01') || row정산월.startsWith('2025-02') || 
            row정산월.startsWith('2025-03') || row정산월.startsWith('2025-04') || 
            row정산월.startsWith('2025-05') || row정산월.startsWith('2025-06') || 
            row정산월.startsWith('2025-07') || row정산월.startsWith('2025-08') || 
            row정산월.startsWith('2025-09') || row정산월.startsWith('2025-10')
          )) {
            sql2025_01_10ExcludedCount++;
            // 디버깅: 처음 5개만 로그
            if (sql2025_01_10ExcludedCount <= 5) {
              console.log(`   ⚠️ [SQL 데이터 제외] 2025-01~2025-10 기간 데이터 제외: 정산월="${row정산월}", 금액=${row.amount || 0}, 거래처명="${row.거래처명 || ''}"`);
            }
            continue; // 2025-01~2025-10 데이터는 제외
          }
          
          sql2025_11_plusCount++;
          
          // 🔥 디버깅: 2025-11 데이터의 merchant 값 확인
          if (row정산월 && row정산월.startsWith('2025-11') && sql2025_11_plusCount <= 5) {
            console.log(`\n🔍 [readExcelAndRespond] 2025-11 데이터 처리:`);
            console.log(`   정산월: "${row정산월}"`);
            console.log(`   row.merchant: "${row.merchant || '(없음)'}" (타입: ${typeof row.merchant})`);
            console.log(`   row.accountName: "${row.accountName || '(없음)'}"`);
            console.log(`   row 객체의 모든 키:`, Object.keys(row).join(', '));
          }
          
          // getSettlementDataFromSQL에서 이미 다음 필드들이 설정되어 있음:
          // - settlementMonth: 정산월
          // - paymentDate: 지급일 (yyyy-mm-dd 형식)
          // - merchant: 사용처 (계정명이 '미지급금_사내'이면 거래처명, 그 외에는 SQL의 "사용처" 컬럼)
          // - amount: 금액
          // - note: 비고
          // - accountName: 계정명
          // - 매칭방법, 매치율: 매칭 정보
          // - isFromSQL: true
          
          const settlementMonthValue = row.settlementMonth || row.정산월 || null;
          const resultItem = {
            month: settlementMonthValue, // 프론트엔드 필터링을 위해 month 필드 추가
            settlementMonth: settlementMonthValue,
            paymentDate: row.paymentDate || null,
            merchant: row.merchant || '', // getSettlementDataFromSQL에서 이미 계산된 값 사용
            amount: row.amount || row.출금액 || 0,
            note: row.note || row.비고 || '',
            accountName: row.accountName || '-',
            isFromSQL: true,
            매칭방법: row.매칭방법 || '알수없음',
            매치율: row.매치율 || 0
          };
          
          // 🔥 디버깅: 최종 resultItem의 merchant 값 확인
          if (row정산월 && row정산월.startsWith('2025-11') && sql2025_11_plusCount <= 5) {
            console.log(`   📋 최종 resultItem.merchant: "${resultItem.merchant || '(없음)'}"`);
          }

          allSettledDetail.push(resultItem);
        }
        
        console.log(`   📊 SQL 데이터 처리 결과:`);
        console.log(`      - 2025-01~2025-10 제외: ${sql2025_01_10ExcludedCount}개 (엑셀 데이터와 중복 방지)`);
        console.log(`      - 2025-11 이후 추가: ${sql2025_11_plusCount}개`);
        console.log(`      - allSettledDetail: ${beforeCount}개 → ${allSettledDetail.length}개`);
        console.log(`   📊 allSettledDetail에 추가: ${beforeCount}개 → ${allSettledDetail.length}개`);
        
        // 🔥 2025-11 데이터 확인
        const sql2025_11 = sqlSettledData.filter(item => {
          const month = item.month || item.settlementMonth || '';
          return month && month.startsWith('2025-11');
        });
        console.log(`   📊 SQL에서 가져온 2025-11 데이터: ${sql2025_11.length}개`);
        if (sql2025_11.length > 0) {
          console.log(`\n${"=".repeat(80)}`);
          console.log(`📊 [2025-11 데이터 ${sql2025_11.length}건 상세 확인]`);
          sql2025_11.forEach((item, idx) => {
            console.log(`\n   ${idx + 1}건:`);
            console.log(`      정산월: "${item.month || item.settlementMonth}"`);
            console.log(`      사용처: "${item.merchant || ''}"`);
            console.log(`      계정명: "${item.accountName || '(없음)'}"`);
            console.log(`      매칭방법: "${item.매칭방법 || '없음'}"`);
            console.log(`      매치율: ${item.매치율 !== undefined ? item.매치율 : '없음'} ${item.매치율 !== undefined ? `(${(item.매치율 * 100).toFixed(1)}%)` : ''}`);
            console.log(`      OpenAI 매칭 여부: ${item.매칭방법 === 'OpenAI매칭' ? '✅ 예' : '❌ 아니오'}`);
            console.log(`      비고: "${(item.note || '').substring(0, 100)}..."`);
          });
          console.log(`${"=".repeat(80)}\n`);
        }
        
        // SQL 데이터의 월별 합계 계산 및 병합 (미정산 데이터 제외)
        sqlSettledData.forEach((item) => {
          const month = item.month || item.settlementMonth || null;
          if (month) {
            // 미정산 데이터 제외 (정산월에 "미정산"이 포함된 경우 제외)
            if (month.includes('미정산') || month.includes('_미정산')) {
              return;
            }
            const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
            if (allMonthlyMap.has(month)) {
              allMonthlyMap.set(month, allMonthlyMap.get(month) + amount);
            } else {
              allMonthlyMap.set(month, amount);
            }
          }
        });
        
        console.log(`📊 정산 데이터 병합 완료: 엑셀(2025-10 이하) ${allSettledDetail.length - sqlSettledData.length}개, SQL(2025-11 이후) ${sqlSettledData.length}개`);
      } else {
        console.log(`⚠️ SQL에서 정산 데이터를 가져오지 못했습니다 (환경 변수 확인 필요 또는 데이터 없음)`);
        console.log(`   💡 가능한 원인:`);
        console.log(`   1. SQL 연결 정보가 설정되지 않음`);
        console.log(`   2. SQL 테이블에 2025-11 이후 데이터가 없음`);
        console.log(`   3. 사용자 필터("${normalizedUserName || '없음'}")에 맞는 데이터가 없음`);
      }

      // 🔥 SQL에서 미정산 데이터 가져오기 (병렬 처리 완료)
      // 미정산 상세내역: [dbo].[ERP_전표상세조회_자금]
      // 🔥 미정산 데이터 조회 (이미 병렬로 실행됨)
      console.log(`\n${"=".repeat(80)}`);
      console.log(`📊 SQL 미정산 데이터 조회 완료 (병렬 처리)`);
      console.log(`   사용자 필터: ${normalizedUserName || '전체'}`);
      console.log(`${"=".repeat(80)}`);
      
      const sqlUnsettledData = sqlUnsettledDetail.status === 'fulfilled' ? sqlUnsettledDetail.value : [];
      
      if (sqlUnsettledDetail.status === 'rejected') {
        console.error(`❌ SQL 미정산 데이터 조회 실패:`, sqlUnsettledDetail.reason);
      }
      
      console.log(`\n📊 SQL 미정산 데이터 조회 결과: ${sqlUnsettledData.length}개 항목`);
      if (sqlUnsettledData.length > 0) {
        console.log(`✅ SQL에서 미정산 ${sqlUnsettledData.length}개 데이터 조회 완료`);
        console.log(`   📋 첫 번째 항목 샘플:`, sqlUnsettledData[0]);
        
        // SQL 미정산 데이터를 allUnsettledData에 병합
        // 🔥 SQL 미정산 데이터 추가 전 계정명 확인
        if (sqlUnsettledData.length > 0) {
          console.log(`\n🔍 SQL 미정산 데이터 추가 전 계정명 확인:`);
          console.log(`   - 총 ${sqlUnsettledData.length}개 항목`);
          sqlUnsettledData.slice(0, 3).forEach((item, idx) => {
            console.log(`   ${idx + 1}. accountName: "${item.accountName || '(없음)'}" (타입: ${typeof item.accountName}), 비고: "${(item.note || '').substring(0, 50)}..."`);
          });
        }
        
        allUnsettledData.push(...sqlUnsettledData);
        
        // SQL 미정산 데이터의 합계 계산 및 병합
        const sqlUnsettledAmount = sqlUnsettledData.reduce((sum, item) => sum + (item.amount || 0), 0);
        allUnsettledAmount += sqlUnsettledAmount;
        
        console.log(`📊 미정산 데이터 병합 완료:`);
        console.log(`   - 엑셀 데이터: ${allUnsettledData.length - sqlUnsettledData.length}개 (제외됨)`);
        console.log(`   - SQL 데이터: ${sqlUnsettledData.length}개`);
        console.log(`   - SQL 합계: ${sqlUnsettledAmount.toLocaleString()}원`);
        console.log(`   - 전체 미정산 합계: ${allUnsettledAmount.toLocaleString()}원`);
      } else {
        console.log(`\n${"=".repeat(80)}`);
        console.log(`⚠️ SQL에서 미정산 데이터를 가져오지 못했습니다`);
        console.log(`   조회 결과: ${sqlUnsettledData.length}개 항목`);
        console.log(`\n   가능한 원인:`);
        console.log(`   1. SQL 연결 정보가 설정되지 않음`);
        console.log(`      → 환경 변수 확인: DB_HOST, DB_USER, DB_PASSWORD, DB_NAME`);
        console.log(`   2. SQL 테이블에 데이터가 없음`);
        console.log(`      → 테이블: ${process.env.DB_TABLE_UNSETTLED || '[dbo].[ERP_전표상세조회_자금]'}`);
        console.log(`   3. 필터 조건에 맞는 데이터가 없음`);
        if (normalizedUserName) {
          console.log(`      → 조건: 반제일 IS NULL AND 사용자 LIKE '%${normalizedUserName}%'`);
        } else {
          console.log(`      → 조건: 없음 (모든 데이터 조회)`);
        }
        console.log(`   4. 사용자 필터에 맞는 데이터가 없음`);
        console.log(`      → 사용자 필터: "${normalizedUserName || '없음 (전체 조회)'}"`);
        if (normalizedUserName) {
          console.log(`      → SQL 쿼리에서 "반제일 IS NULL AND 사용자 LIKE '%${normalizedUserName}%'" 조건 적용됨`);
        }
        console.log(`${"=".repeat(80)}\n`);
      }

    // 🔥 최종 안전장치: 모든 파일 병합 후에도 userName 기준으로 한 번 더 전체 필터링
      let finalSettledDetail = allSettledDetail;
      let finalUnsettledData = allUnsettledData;
      let finalUnsettledAmount = allUnsettledAmount;
      let finalMonthlyMap = allMonthlyMap;

      if (normalizedUserName) {
        console.log(`\n🔍 최종 사용자 필터링 적용: ${normalizedUserName}`);
        const beforeSettled = allSettledDetail.length;
        const beforeUnsettled = allUnsettledData.length;
        
        // 🔥 필터링 전 SQL 데이터 확인
        const sqlSettledBeforeFilter = allSettledDetail.filter(item => item.isFromSQL);
        console.log(`   📊 필터링 전 정산 데이터: 전체 ${beforeSettled}개 (SQL: ${sqlSettledBeforeFilter.length}개, 엑셀: ${beforeSettled - sqlSettledBeforeFilter.length}개)`);
        const sql2025_11_before = sqlSettledBeforeFilter.filter(item => {
          const month = item.month || item.settlementMonth || '';
          return month && month.startsWith('2025-11');
        });
        console.log(`   📊 필터링 전 2025-11 SQL 데이터: ${sql2025_11_before.length}개`);
        // 성능 최적화: 상세 로그 제거

        // 🔥 최종 필터링: 엑셀 데이터는 이미 첫 번째 필터링(1849번 줄)에서 A열 거래처명으로 필터링되었으므로 통과
        // SQL 데이터는 이미 SQL 쿼리에서 필터링되었으므로 통과
        let 디버그카운트최종 = 0;
        finalSettledDetail = allSettledDetail.filter(item => {
          // 🔥 SQL 데이터는 이미 SQL 쿼리에서 필터링되었으므로 통과
          if (item.isFromSQL) {
            // 성능 최적화: 디버깅 로그 제거
            return true; // SQL 데이터는 필터링 없이 통과
          }
          
          // 🔥 엑셀 데이터는 이미 첫 번째 필터링(1849번 줄)에서 A열 거래처명으로 필터링되었으므로 통과
          // 성능 최적화: 디버깅 로그 제거
          
          return true; // 엑셀 데이터는 이미 필터링되었으므로 통과
        });

        // 🔥 필터링 후 SQL 데이터 확인
        const sqlSettledCount = finalSettledDetail.filter(item => item.isFromSQL).length;
        console.log(`   📊 필터링 후 정산 데이터: 전체 ${finalSettledDetail.length}개 (SQL: ${sqlSettledCount}개, 엑셀: ${finalSettledDetail.length - sqlSettledCount}개)`);
        
        // 🔥 2025-11 데이터 확인
        const settled2025_11 = finalSettledDetail.filter(item => {
          const month = item.month || item.settlementMonth || '';
          return month && month.startsWith('2025-11');
        });
        console.log(`   📊 2025-11 정산 데이터: ${settled2025_11.length}개`);
        if (settled2025_11.length > 0) {
          console.log(`   📋 2025-11 첫 번째 항목:`, {
            정산월: settled2025_11[0].month || settled2025_11[0].settlementMonth,
            사용처: settled2025_11[0].merchant,
            계정명: settled2025_11[0].accountName || '(없음)',
            금액: settled2025_11[0].amount,
            isFromSQL: settled2025_11[0].isFromSQL
          });
        }

        console.log(`   📊 필터링 전 미정산 데이터: ${allUnsettledData.length}개`);
        if (allUnsettledData.length > 0) {
          console.log(`   📋 미정산 데이터 샘플 (처음 3개):`, allUnsettledData.slice(0, 3).map(item => ({
            정산월: item.settlementMonth || item.month,
            사용처: item.merchant,
            계정명: item.accountName || '(없음)',
            계정명타입: typeof item.accountName,
            금액: item.amount
          })));
        }
        finalUnsettledData = allUnsettledData.filter((item, index) => {
          // 🔥 SQL에서 가져온 데이터는 이미 사용자 필터링이 적용되었으므로 건너뛰기
          if (item.isFromSQL) {
            return true; // SQL 데이터는 필터링 없이 통과
          }
          
          // 엑셀 데이터만 merchant로 필터링
          const 거래처명값 = item.merchant || "";
          const 매칭결과 = matchUserByMerchant(거래처명값, normalizedUserName);
          
          // 디버깅: 처음 10개 항목만 로그
          if (index < 10) {
            console.log(`   [미정산필터] index=${index}, isFromSQL=${item.isFromSQL || false}, merchant="${거래처명값}", 매칭=${매칭결과}`);
          }
          
          return 매칭결과;
        });
        console.log(`   📊 필터링 후 미정산 데이터: ${finalUnsettledData.length}개`);
        
        // 🔥 필터링 후 accountName 확인
        if (finalUnsettledData.length > 0) {
          console.log(`\n🔍 필터링 후 accountName 확인:`);
          finalUnsettledData.slice(0, 3).forEach((item, idx) => {
            const hasAccountName = 'accountName' in item;
            console.log(`   ${idx + 1}. accountName 필드 존재: ${hasAccountName}, 값: "${item.accountName || '(없음)'}" (타입: ${typeof item.accountName})`);
            console.log(`      전체 객체 키: ${Object.keys(item).join(', ')}`);
            console.log(`      비고: "${(item.note || '').substring(0, 50)}..."`);
          });
        }
        
        if (finalUnsettledData.length === 0 && allUnsettledData.length > 0) {
          console.log(`   ⚠️ 경고: 필터링으로 인해 모든 미정산 데이터가 제외되었습니다!`);
          console.log(`   💡 사용자 필터("${normalizedUserName}")와 merchant 값이 일치하지 않습니다.`);
          console.log(`   💡 merchant 값 샘플:`, allUnsettledData.slice(0, 5).map(item => item.merchant));
        }

      finalUnsettledAmount = finalUnsettledData.reduce(
        (sum, item) => sum + (item.amount || 0),
        0
      );

      // 월별 합계 재계산 (필터링된 정산 데이터 기준, 미정산 데이터 제외)
      // 🔥 정산월(month) 필드 기준으로만 집계 (N열 정산월 값 사용)
      finalMonthlyMap = new Map();
      if (Array.isArray(finalSettledDetail)) {
        let monthNullCount = 0;
        let monthEmptyCount = 0;
        let monthValidCount = 0;
        
        finalSettledDetail.forEach((item, idx) => {
          try {
            // 🔥 정산월은 item.month 필드를 우선 사용 (N열에서 읽은 값)
            const month = item.month || item.settlementMonth || null;
            
            if (!month) {
              monthNullCount++;
              // 디버깅: 처음 5개만 로그
              if (idx < 5) {
                console.warn(`   ⚠️ [월별집계] index=${idx}: 정산월 없음, item.month="${item.month}", item.settlementMonth="${item.settlementMonth}", amount=${item.amount}`);
              }
              return;
            }
            
            // 빈 문자열 체크
            if (String(month).trim() === '') {
              monthEmptyCount++;
              if (idx < 5) {
                console.warn(`   ⚠️ [월별집계] index=${idx}: 정산월 빈 문자열, amount=${item.amount}`);
              }
              return;
            }
            
            // 미정산 데이터 제외 (정산월에 "미정산"이 포함된 경우 제외)
            if (month.includes('미정산') || month.includes('_미정산')) {
              return;
            }
            
            monthValidCount++;
            const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
            
            if (finalMonthlyMap.has(month)) {
              finalMonthlyMap.set(month, finalMonthlyMap.get(month) + amount);
            } else {
              finalMonthlyMap.set(month, amount);
            }
            
            // 디버깅: 처음 10개만 로그
            if (idx < 10) {
              console.log(`   [월별집계] index=${idx}: 정산월="${month}", 금액=${amount}, 누적합계=${finalMonthlyMap.get(month)}`);
            }
          } catch (e) {
            console.error(`⚠️ [월별집계] index=${idx} 오류:`, e, item);
          }
        });
        
        console.log(`\n📊 월별 집계 통계:`);
        console.log(`   ✅ 정산월 있음: ${monthValidCount}개`);
        console.log(`   ⚠️ 정산월 없음: ${monthNullCount}개`);
        console.log(`   ⚠️ 정산월 빈 문자열: ${monthEmptyCount}개`);
        console.log(`   📋 월별 집계 결과: ${finalMonthlyMap.size}개 월`);
        finalMonthlyMap.forEach((amount, month) => {
          console.log(`      ${month}: ${amount.toLocaleString()}원`);
        });
      }

      // 🔥 최종 필터링된 정산 데이터 합계 계산
      const finalSettledTotal = Array.isArray(finalSettledDetail) ? finalSettledDetail.reduce((sum, item) => {
        try {
          const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
          return sum + amount;
        } catch (e) {
          console.error('⚠️ 최종 합계 계산 중 오류:', e, item);
          return sum;
        }
      }, 0) : 0;
      
      console.log(
        `   ▶ 최종 사용자 필터 결과 - 정산: ${beforeSettled} → ${finalSettledDetail.length}, 미정산: ${beforeUnsettled} → ${finalUnsettledData.length}`
      );
      console.log(`   ▶ 필터링 전 전체 정산 합계: ${beforeFilterTotal.toLocaleString()}원 (${allSettledDetail.length}개 항목)`);
      console.log(`   ▶ 필터링 후 최종 정산 합계: ${finalSettledTotal.toLocaleString()}원 (${finalSettledDetail.length}개 항목)`);
      console.log(`   ▶ 최종 미정산 합계: ${finalUnsettledAmount}`);
    }

    // 월별 정산 요약 생성 (병합된 데이터)
    const monthly = Array.from(finalMonthlyMap.entries())
      .map(([month, amount]) => ({ month, amount }))
      .sort((a, b) => (a.month || '').localeCompare(b.month || ''));

    // 🔥 최종 응답 데이터 생성 (병합된 데이터 사용)
    responseData = {
      success: true,
      code: 200,
      message: '개인정산 데이터 조회 성공',
      data: {
        settled: {
          monthly: isUnsettledSheet ? [] : monthly,
          detail: isUnsettledSheet ? [] : finalSettledDetail.sort((a, b) => {
            const aMonth = a.month || a.settlementMonth || '';
            const bMonth = b.month || b.settlementMonth || '';
            if (aMonth !== bMonth) {
              return aMonth.localeCompare(bMonth);
            }
            return (a.date || '').localeCompare(b.date || '');
          })
        },
        unsettled: {
          amount: finalUnsettledAmount,
          detail: finalUnsettledData
        }
      }
    };

    console.log("\n" + "=".repeat(80));
    console.log("💾 responseData에 데이터 저장 완료");
    console.log(`   📊 정산 상세 내역: ${responseData.data.settled.detail.length}개`);
    console.log(`   📊 미정산 상세 내역: ${responseData.data.unsettled.detail.length}개`);
    console.log(`   💰 미정산 합계: ${finalUnsettledAmount.toLocaleString()}원`);
    
    // 🔥 2025-11 데이터 최종 확인
    const final2025_11 = responseData.data.settled.detail.filter(item => {
      const month = item.month || item.settlementMonth || '';
      return month && month.startsWith('2025-11');
    });
    console.log(`   📊 최종 응답 데이터 중 2025-11 데이터: ${final2025_11.length}개`);
    if (final2025_11.length > 0) {
      final2025_11.slice(0, 3).forEach((item, idx) => {
        console.log(`      ${idx + 1}. 정산월: "${item.month || item.settlementMonth}", merchant: "${item.merchant}", isFromSQL: ${item.isFromSQL || false}`);
        console.log(`         accountName: "${item.accountName || '(없음)'}" (타입: ${typeof item.accountName})`);
        console.log(`         accountName 필드 존재: ${'accountName' in item}`);
        console.log(`         비고: "${(item.note || '').substring(0, 50)}..."`);
      });
    } else {
      console.log(`   ⚠️ 최종 응답에 2025-11 데이터가 없습니다!`);
      const sql2025_11_in_final = finalSettledDetail.filter(item => {
        const month = item.month || item.settlementMonth || '';
        return month && month.startsWith('2025-11');
      });
      console.log(`   💡 finalSettledDetail에는 2025-11 데이터가 ${sql2025_11_in_final.length}개 있습니다.`);
    }
    console.log(`\n   📋 데이터 흐름 확인:`);
    console.log(`   - allUnsettledData (SQL 조회 후): ${allUnsettledData.length}개`);
    console.log(`   - finalUnsettledData (필터링 후): ${finalUnsettledData.length}개`);
    console.log(`   - responseData.data.unsettled.detail (최종 응답): ${responseData.data.unsettled.detail.length}개`);
    
    // 🔥 2025-01~08 정산월 데이터의 accountName 확인
    const settled2025_01_08 = responseData.data.settled.detail.filter(item => {
      const month = item.month || item.settlementMonth || '';
      return month && (month.startsWith('2025-01') || month.startsWith('2025-02') || month.startsWith('2025-03') || month.startsWith('2025-04') || month.startsWith('2025-05') || month.startsWith('2025-06') || month.startsWith('2025-07') || month.startsWith('2025-08'));
    });
    if (settled2025_01_08.length > 0) {
      console.log(`\n🔍 2025-01~08 정산월 데이터 accountName 확인:`);
      console.log(`   총 ${settled2025_01_08.length}개 항목`);
      const 빈계정명개수 = settled2025_01_08.filter(item => !item.accountName || item.accountName.trim() === '').length;
      console.log(`   ⚠️ 계정명이 비어있는 항목: ${빈계정명개수}개`);
      settled2025_01_08.slice(0, 10).forEach((item, idx) => {
        console.log(`   ${idx + 1}. 정산월: "${item.month || item.settlementMonth}", accountName: "${item.accountName || '(없음)'}" ${!item.accountName || item.accountName.trim() === '' ? '⚠️' : '✅'}`);
      });
    }
    
    // 🔥 2025-09 정산월 데이터의 accountName 확인 (비교용)
    const settled2025_09 = responseData.data.settled.detail.filter(item => {
      const month = item.month || item.settlementMonth || '';
      return month && month.startsWith('2025-09');
    });
    if (settled2025_09.length > 0) {
      console.log(`\n🔍 2025-09 정산월 데이터 accountName 확인 (비교용):`);
      console.log(`   총 ${settled2025_09.length}개 항목`);
      settled2025_09.slice(0, 3).forEach((item, idx) => {
        console.log(`   ${idx + 1}. 정산월: "${item.month || item.settlementMonth}", accountName: "${item.accountName || '(없음)'}"`);
      });
    }
    
    if (responseData.data.unsettled.detail.length > 0) {
      console.log(`\n   ✅ 미정산 데이터가 응답에 포함되었습니다!`);
      const firstItem = responseData.data.unsettled.detail[0];
      console.log(`   📋 미정산 첫 번째 항목 샘플:`, {
        정산월: firstItem.settlementMonth || firstItem.month,
        사용처: firstItem.merchant,
        계정명: firstItem.accountName || '(없음)',
        계정명타입: typeof firstItem.accountName,
        계정명값: JSON.stringify(firstItem.accountName),
        금액: firstItem.amount,
        비고: (firstItem.note || '').substring(0, 50) + '...'
      });
      
      // 🔥 계정명이 "-"인 항목들 확인
      const 계정명하이픈항목 = responseData.data.unsettled.detail.filter(item => 
        item.accountName === '-' || item.accountName === '' || !item.accountName
      );
      if (계정명하이픈항목.length > 0) {
        console.log(`\n   ⚠️ 계정명이 "-"인 항목: ${계정명하이픈항목.length}개`);
        계정명하이픈항목.slice(0, 3).forEach((item, idx) => {
          console.log(`      ${idx + 1}. 비고: "${(item.note || '').substring(0, 50)}...", 계정명: "${item.accountName || '(없음)'}"`);
        });
      }
      
      // 계정명이 없는 항목 확인
      const 계정명없는응답항목 = responseData.data.unsettled.detail.filter(item => !item.accountName || item.accountName === '' || item.accountName === '-');
      if (계정명없는응답항목.length > 0) {
        console.log(`   ⚠️ 응답 데이터 중 계정명이 없는 항목: ${계정명없는응답항목.length}개 / 전체 ${responseData.data.unsettled.detail.length}개`);
      } else {
        console.log(`   ✅ 응답 데이터의 모든 항목에 계정명이 있습니다.`);
      }
    } else {
      console.log(`\n   ⚠️ 미정산 데이터가 응답에 포함되지 않았습니다!`);
      if (allUnsettledData.length > 0) {
        console.log(`   💡 SQL에서 ${allUnsettledData.length}개 데이터를 가져왔지만 필터링에서 모두 제외되었습니다.`);
        console.log(`   💡 사용자 필터("${normalizedUserName || '없음'}")를 확인하세요.`);
      } else {
        console.log(`   💡 SQL에서 데이터를 가져오지 못했습니다.`);
        console.log(`   💡 서버 콘솔의 SQL 조회 로그를 확인하세요.`);
      }
    }
    console.log("=".repeat(80) + "\n");
    console.log("📝 각 법인별 result 파일은 이미 생성되었습니다 (병합하지 않음)");


    // 🔥 최종 응답 전 데이터 확인 로그
    console.log(`\n${"=".repeat(80)}`);
    console.log(`📤 최종 응답 데이터 확인 (readExcelAndRespond)`);
    console.log(`   responseData.success: ${responseData.success}`);
    console.log(`   responseData.data.settled.detail.length: ${responseData.data.settled.detail.length}`);
    console.log(`   responseData.data.unsettled.amount: ${responseData.data.unsettled.amount}`);
    console.log(`   responseData.data.unsettled.detail.length: ${responseData.data.unsettled.detail.length}`);
    
    // 🔥 정산 데이터 accountName 확인 (특히 2025-11 SQL 데이터)
    if (responseData.data.settled.detail.length > 0) {
      console.log(`   ✅ 정산 데이터 ${responseData.data.settled.detail.length}개가 응답에 포함됨`);
      const sql2025_11 = responseData.data.settled.detail.filter(item => {
        const month = item.month || item.settlementMonth || '';
        return item.isFromSQL && month && month.startsWith('2025-11');
      });
      if (sql2025_11.length > 0) {
        console.log(`   📊 2025-11 SQL 정산 데이터: ${sql2025_11.length}개`);
        sql2025_11.slice(0, 3).forEach((item, idx) => {
          console.log(`      ${idx + 1}. 정산월: "${item.month || item.settlementMonth}", accountName: "${item.accountName || '(없음)'}"`);
          console.log(`         accountName 필드 존재: ${'accountName' in item}`);
          console.log(`         전체 객체 키: ${Object.keys(item).join(', ')}`);
        });
      }
    } else {
      console.log(`   ⚠️ 정산 데이터가 응답에 포함되지 않음 (0개)`);
    }
    
    if (responseData.data.unsettled.detail.length > 0) {
      console.log(`   ✅ 미정산 데이터 ${responseData.data.unsettled.detail.length}개가 응답에 포함됨`);
      console.log(`   📋 첫 번째 항목:`, {
        정산월: responseData.data.unsettled.detail[0].settlementMonth || responseData.data.unsettled.detail[0].month,
        사용처: responseData.data.unsettled.detail[0].merchant,
        계정명: responseData.data.unsettled.detail[0].accountName,
        비고: (responseData.data.unsettled.detail[0].note || '').substring(0, 50) + '...',
        금액: responseData.data.unsettled.detail[0].amount
      });
    } else {
      console.log(`   ⚠️ 미정산 데이터가 응답에 포함되지 않음 (0개)`);
    }
    console.log(`${"=".repeat(80)}\n`);
    
    // 🔥 응답 전 최종 확인: accountName이 실제로 포함되어 있는지 검증 (정산 + 미정산)
    if (responseData.data.settled.detail.length > 0) {
      console.log(`\n${"=".repeat(80)}`);
      console.log(`🔍 응답 전 최종 검증: 정산 데이터 accountName 필드 확인`);
      const sql2025_11 = responseData.data.settled.detail.filter(item => {
        const month = item.month || item.settlementMonth || '';
        return item.isFromSQL && month && month.startsWith('2025-11');
      });
      if (sql2025_11.length > 0) {
        sql2025_11.slice(0, 3).forEach((item, idx) => {
          const hasAccountName = 'accountName' in item;
          const accountNameValue = item.accountName;
          console.log(`   ${idx + 1}. accountName 필드 존재: ${hasAccountName}, 값: "${accountNameValue}" (타입: ${typeof accountNameValue})`);
          console.log(`      전체 객체 키: ${Object.keys(item).join(', ')}`);
          console.log(`      비고: "${(item.note || '').substring(0, 50)}..."`);
        });
      } else {
        console.log(`   ⚠️ 2025-11 SQL 정산 데이터가 없습니다.`);
      }
      console.log(`${"=".repeat(80)}\n`);
    }
    
    if (responseData.data.unsettled.detail.length > 0) {
      console.log(`\n${"=".repeat(80)}`);
      console.log(`🔍 응답 전 최종 검증: 미정산 데이터 accountName 필드 확인`);
      responseData.data.unsettled.detail.slice(0, 3).forEach((item, idx) => {
        const hasAccountName = 'accountName' in item;
        const accountNameValue = item.accountName;
        console.log(`   ${idx + 1}. accountName 필드 존재: ${hasAccountName}, 값: "${accountNameValue}" (타입: ${typeof accountNameValue})`);
        console.log(`      전체 객체 키: ${Object.keys(item).join(', ')}`);
        console.log(`      비고: "${(item.note || '').substring(0, 50)}..."`);
      });
      console.log(`${"=".repeat(80)}\n`);
    }

    res.json(responseData);

    // 캐시 저장
    try {
      responseCache.set(cacheKey, { data: responseData, timestamp: Date.now() });
    } catch (err) {
      console.error('⚠️ 캐시 저장 중 오류:', err);
    }

  } catch (error) {
    console.error("❌ 개인정산 데이터 로드 오류:", error);
    res.status(500).json({ 
      success: false,
      error: error.message || "서버 오류" 
    });
  }
}

// ===================================================
// 📌 개인정산 데이터 로드 API
// match_data3_result.xlsx 파일에서 데이터 가져오기
// ===================================================
app.get('/api/ping', (req, res) => {
  res.json({ status: 'ok' });
});

app.get("/api/personal-settlement", async (req, res) => {
  const sheetName = "2025"; // 정산 데이터 시트
  const userName = req.query.username || req.query.userName || null;
  const period = req.query.period || null; // 조회 기간 (예: "2025-01 ~ 2025-12")
  console.log("\n" + "=".repeat(80));
  console.log("🔥 /api/personal-settlement 요청 받음!");
  console.log(`   시트명: ${sheetName}`);
  console.log(`   사용자명: ${userName || '전체'}`);
  console.log(`   조회 기간: ${period || '없음'}`);
  console.log(`   쿼리 파라미터:`, req.query);
  console.log("=".repeat(80) + "\n");
  return readExcelAndRespond(res, sheetName, userName, period);
});


// ===================================================
// 📌 프론트엔드에서 호출하는 /api/all-data
// ⚠️ 반드시 /api/personal-settlement 호출 후에만 사용 가능
// 여기서 match_data3_result 를 넘겨준다!
// ===================================================
app.get("/api/all-data", (req, res) => {
  try {
    console.log("📌 /api/all-data 호출됨");
    console.log("📌 /api/all-data 쿼리:", req.query);

    // 1️⃣ 순서 확인: /api/personal-settlement가 먼저 호출되어야 함
    if (!responseData) {
      console.error("❌ /api/all-data 오류: responseData가 없습니다. 먼저 /api/personal-settlement를 호출해야 합니다.");
      return res.status(400).json({ 
        success: false,
        error: "정산 데이터가 없습니다. 먼저 /api/personal-settlement를 호출해주세요." 
      });
    }

    console.log("✅ /api/all-data: 저장된 데이터 반환 (총 " + responseData.data.settled.detail.length + "개 항목)");

    res.json({
      success: true,
      match_data3_result: responseData.data.settled.detail,
    });

  } catch (error) {
    console.error("❌ /api/all-data 오류:", error);
    res.status(500).json({ 
      success: false,
      error: "데이터 로딩 오류: " + error.message 
    });
  }
});

// ===================================================
// 📌 관리자 화면: 특정 시트 데이터 조회 API
// ===================================================
app.get("/api/data/:sheetName", async (req, res) => {
  try {
    const sheetName = req.params.sheetName;
    const page = Number(req.query.page) || 1;
    const limit = Number(req.query.limit) || 100;
    // MOCA 파일만 처리 (기존법인 파일 제외)
    const mocaFile = ADDITIONAL_EXCEL_FILES.find(f => f.includes("moca")) || "./match_data_moca.xlsx";
    const excelFilePath = req.query.excelFile || mocaFile;
    const excelPath = getExcelFilePath(excelFilePath);

    console.log(`📋 /api/data/${sheetName} 호출 - page: ${page}, limit: ${limit}`);

    // 파일 존재 확인
    if (!fs.existsSync(excelPath)) {
      console.warn(`⚠️ MOCA 파일이 존재하지 않습니다: ${excelPath}`);
      return res.json({
        success: true,
        data: [],
        headers: [],
        totalRows: 0,
        page,
        limit,
        totalPages: 0,
      });
    }

    const { data, headers, totalRows, totalPages } = await getExcelData(
      excelPath,
      sheetName,
      null,
      page,
      limit
    );

    res.json({
      success: true,
      data,
      headers,
      totalRows,
      page,
      limit,
      totalPages,
    });
  } catch (error) {
    console.error("❌ /api/data/:sheetName 오류:", error);
    // 파일이 없거나 읽을 수 없는 경우 빈 데이터 반환 (오류로 처리하지 않음)
    res.json({
      success: true,
      data: [],
      headers: [],
      totalRows: 0,
      page: Number(req.query.page) || 1,
      limit: Number(req.query.limit) || 100,
      totalPages: 0,
    });
  }
});

app.get("/api/sheets", async (req, res) => {
  try {
    // MOCA 파일만 처리 (기존법인 파일 제외)
    const mocaFile = ADDITIONAL_EXCEL_FILES.find(f => f.includes("moca")) || "./match_data_moca.xlsx";
    const excelPath = getExcelFilePath(mocaFile);
    
    // 파일 존재 확인
    if (!fs.existsSync(excelPath)) {
      console.warn(`⚠️ MOCA 파일이 존재하지 않습니다: ${excelPath}`);
      return res.json({
        success: true,
        sheets: []
      });
    }
    
    const sheets = await getSheetNames(excelPath);
    res.json({
      success: true,
      sheets
    });
  } catch (error) {
    console.error("❌ /api/sheets 오류:", error);
    // 파일이 없거나 읽을 수 없는 경우 빈 배열 반환 (오류로 처리하지 않음)
    res.json({
      success: true,
      sheets: []
    });
  }
});

// ===================================================
// 📌 캐시 무효화 API (엑셀 파일 변경 시 호출)
// ===================================================
app.post("/api/clear-cache", (req, res) => {
  try {
    console.log("🔄 캐시 무효화 요청 받음");
    responseData = null;
    console.log("✅ 서버 캐시가 무효화되었습니다. 다음 요청 시 최신 데이터가 로드됩니다.");
    res.json({
      success: true,
      message: "캐시가 무효화되었습니다.",
    });
  } catch (error) {
    console.error("❌ 캐시 무효화 오류:", error);
    res.status(500).json({
      success: false,
      error: error.message || "캐시 무효화 실패",
    });
  }
});


// ===================================================
// 🚀 정적 파일 서빙 (API 라우트 이후에 설정, /api 경로 제외)
// ===================================================
app.use((req, res, next) => {
  // /api로 시작하는 경로는 정적 파일 미들웨어를 건너뜀
  if (req.path.startsWith('/api')) {
    return next();
  }
  express.static(path.join(__dirname, "public"))(req, res, next);
});

// ===================================================
// 🚀 루트 라우트 (모든 라우트 이후에 설정)
// ===================================================
app.get("/", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "index.html"));
});

// 🔥 미정산 데이터 API 추가
app.get("/api/unsettled-data", async (req, res) => {
  try {
    const userName = req.query.username || req.query.userName || null;
    console.log(`\n${"=".repeat(80)}`);
    console.log(`🔥 /api/unsettled-data 엔드포인트 호출됨`);
    console.log(`   쿼리 파라미터:`, req.query);
    console.log(`   사용자명: ${userName || '없음 (전체 조회)'}`);
    console.log(`   요청 시간: ${new Date().toISOString()}`);
    console.log(`${"=".repeat(80)}\n`);
    
    // SQL 연결 정보 확인 (디버깅용)
    console.log(`\n🔍 SQL 연결 정보 확인:`);
    console.log(`   DB_HOST: ${process.env.DB_HOST || process.env.DB_SERVER || '없음'}`);
    console.log(`   DB_PORT: ${process.env.DB_PORT || '1433 (기본값)'}`);
    console.log(`   DB_USER: ${process.env.DB_USER ? '설정됨' : '없음'}`);
    console.log(`   DB_PASSWORD: ${process.env.DB_PASSWORD ? '설정됨' : '없음'}`);
    console.log(`   DB_NAME: ${process.env.DB_NAME || process.env.DB_DATABASE || '없음'}`);
    console.log(`   DB_TABLE_UNSETTLED: ${process.env.DB_TABLE_UNSETTLED || '[dbo].[ERP_전표상세조회_자금] (기본값)'}`);
    console.log(``);
    
    await readExcelAndRespond(res, "2025_미정산", userName);
  } catch (error) {
    console.error(`\n❌ /api/unsettled-data 엔드포인트 오류:`, error);
    console.error(`   오류 메시지:`, error.message);
    console.error(`   스택:`, error.stack);
    res.status(500).json({
      success: false,
      error: error.message || "미정산 데이터 조회 오류"
    });
  }
});

// 🔥 서버 헬스체크 엔드포인트
app.get("/api/health", (req, res) => {
  res.json({
    success: true,
    status: "running",
    message: "서버가 정상적으로 실행 중입니다."
  });
});

// 🔥 루트 경로 헬스체크 엔드포인트
app.get("/health", (req, res) => {
  res.json({ status: "ok" });
});

// 🔥 서버 자동 시작 API (Windows에서 배치 파일 실행)
app.post("/api/start-server", (req, res) => {
  try {
    const startScriptPath = path.join(__dirname, "start-all.cmd");
    
    // 파일 존재 확인
    if (!fs.existsSync(startScriptPath)) {
      return res.status(404).json({
        success: false,
        error: "start-all.cmd 파일을 찾을 수 없습니다."
      });
    }
    
    // Windows에서 배치 파일 실행 (새 창에서 실행)
    // start 명령어는 비동기로 실행되므로 즉시 응답 반환
    const command = `start "" "${startScriptPath}" --auto`;
    
    exec(command, { 
      cwd: __dirname,
      windowsHide: false // 창이 보이도록 설정
    }, (error, stdout, stderr) => {
      // start 명령어는 즉시 반환되므로 error가 발생해도 정상일 수 있음
      if (error && !error.message.includes('start')) {
        console.error("서버 시작 오류:", error);
        return res.status(500).json({
          success: false,
          error: "서버 시작에 실패했습니다: " + error.message
        });
      }
      
      console.log("서버 시작 명령 실행됨");
      // 성공 응답은 즉시 반환 (서버가 시작되는 동안 대기하지 않음)
      res.json({
        success: true,
        message: "서버 시작 명령이 실행되었습니다. 잠시 후 서버가 시작됩니다."
      });
    });
    
    // exec가 비동기이므로 즉시 응답을 반환하지 않고 위의 콜백에서 처리
    // 하지만 start 명령어는 즉시 반환되므로 타임아웃을 설정하여 안전하게 처리
    setTimeout(() => {
      if (!res.headersSent) {
        res.json({
          success: true,
          message: "서버 시작 명령이 실행되었습니다. 잠시 후 서버가 시작됩니다."
        });
      }
    }, 1000);
    
  } catch (error) {
    console.error("서버 시작 API 오류:", error);
    if (!res.headersSent) {
      res.status(500).json({
        success: false,
        error: error.message || "서버 시작 중 오류가 발생했습니다."
      });
    }
  }
});

// 네트워크 IP 주소 가져오기
function getNetworkIP() {
  const interfaces = os.networkInterfaces();
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      // IPv4이고 내부 주소가 아닌 경우
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return 'localhost';
}

const HOST = '0.0.0.0'; // 모든 네트워크 인터페이스에서 접근 가능
const networkIP = getNetworkIP();

// 캐시 변수 선언
let cachedExcelData = null;
let lastLoadedTime = null;

// 📌 캐시 초기화 API (항상 성공 처리)
app.get('/api/clear-cache', (req, res) => {
    try {
        cachedExcelData = null;
        lastLoadedTime = null;

        console.log('📁 엑셀 캐시가 초기화되었습니다.');
        
        res.json({
            success: true,
            message: '엑셀 캐시가 초기화되었습니다.'
        });
    } catch (error) {
        console.error('⚠️ 캐시 초기화 중 오류:', error);

        // 오류가 발생해도 캐시는 이미 null 상태이므로 실제 문제 없음
        res.json({
            success: true,
            message: '캐시가 초기화되었습니다. (오류 무시)'
        });
    }
});

// 🔥 SQL 미정산 계정명 캐시 무효화 API
// SQL 데이터의 비고가 수정되거나 추가되었을 때 호출하여 캐시를 초기화
app.post('/api/clear-unsettled-account-cache', (req, res) => {
    try {
        const beforeSize = unsettledAccountNameCache.size;
        unsettledAccountNameCache.clear();
        
        console.log(`🔄 SQL 미정산 계정명 캐시 무효화: ${beforeSize}개 항목 삭제`);
        
        res.json({
            success: true,
            message: `SQL 미정산 계정명 캐시가 초기화되었습니다. (${beforeSize}개 항목 삭제)`,
            clearedCount: beforeSize
        });
    } catch (error) {
        console.error('⚠️ SQL 미정산 계정명 캐시 초기화 중 오류:', error);
        res.status(500).json({
            success: false,
            error: error.message || '캐시 초기화 실패'
        });
    }
});

// OpenAI 초기화 제거됨
// const openai = new OpenAI({
//   apiKey: process.env.OPENAI_API_KEY,
// });

// OpenAI 요약 API 제거됨
// app.post("/api/summary", async (req, res) => {
//   ... (OpenAI 요약 로직 제거)
// });

// OpenAI AI 요약 API 제거됨
// app.post("/api/ai-summary", async (req, res) => {
//   ... (OpenAI 요약 로직 제거)
// });

// 🔥 전역 에러 핸들러 추가 (서버가 끊기는 것을 방지)
process.on('uncaughtException', (error) => {
  console.error('❌ 처리되지 않은 예외 발생:', error);
  console.error('스택:', error.stack);
  // 서버를 종료하지 않고 계속 실행
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ 처리되지 않은 Promise 거부:', reason);
  console.error('Promise:', promise);
  if (reason instanceof Error) {
    console.error('에러 스택:', reason.stack);
  }
  // 서버를 종료하지 않고 계속 실행
});

// Express 에러 핸들러 미들웨어 (모든 라우트 정의 후에 추가)
app.use((err, req, res, next) => {
  console.error('❌ Express 미들웨어 에러:', err);
  console.error('요청 경로:', req.path);
  console.error('요청 메서드:', req.method);
  if (err.stack) {
    console.error('에러 스택:', err.stack);
  }
  
  // 응답이 아직 전송되지 않았을 때만 에러 응답 전송
  if (!res.headersSent) {
    res.status(500).json({
      success: false,
      error: err.message || '서버 내부 오류가 발생했습니다.',
      path: req.path
    });
  }
});

// 🔥 서버 시작 - 모든 라우트 정의 후에 호출
const server = app.listen(PORT, HOST, () => {
  console.log(`\n${"=".repeat(80)}`);
  console.log(`🚀 서버가 실행 중입니다. (PID: ${process.pid})`);
  console.log(`   시작 시간: ${new Date().toLocaleString('ko-KR')}`);
  console.log(`${"=".repeat(80)}\n`);
  console.log(`📍 접속 주소:`);
  console.log(`   - 로컬: http://localhost:${PORT}`);
  console.log(`   - 네트워크: http://${networkIP}:${PORT}`);
  console.log(`\n💡 다른 사람과 공유하려면 네트워크 주소를 사용하세요!`);
  console.log(`   같은 네트워크에 연결된 다른 기기에서 접속 가능합니다.`);
  console.log(`\n📁 엑셀 파일 설정:`);
  console.log(`   - MOCA 파일만 처리 (기존법인 파일 제외)`);
  console.log(`   - MOCA 파일: ${ADDITIONAL_EXCEL_FILES.filter(f => f.includes("moca")).join(', ') || '없음'}`);
  console.log(`   - 기본 시트명: ${EXCEL_SHEET_NAME}`);
  console.log(`   - 환경 변수 ADDITIONAL_EXCEL_FILES로 MOCA 파일 경로 변경 가능`);
  console.log(`   - 환경 변수 EXCEL_SHEET_NAME로 시트명 변경 가능`);
  console.log(`\n🔍 모든 요청이 로깅됩니다. API 호출 시 터미널에 로그가 표시됩니다.\n`);
});

// 서버 에러 핸들러 추가
server.on('error', (error) => {
  if (error.code === 'EADDRINUSE') {
    console.error(`❌ 포트 ${PORT}가 이미 사용 중입니다.`);
    console.error(`   다른 프로세스가 포트를 사용하고 있거나 서버가 이미 실행 중일 수 있습니다.`);
    console.error(`   해결 방법:`);
    console.error(`   1. 기존 서버 프로세스를 종료하세요 (PID: ${process.pid})`);
    console.error(`   2. 다른 포트를 사용하세요 (환경 변수 PORT 설정)`);
  } else {
    console.error('❌ 서버 에러 발생:', error);
    console.error('스택:', error.stack);
  }
});

// 서버 연결 종료 핸들러
server.on('close', () => {
  console.log('⚠️ 서버 연결이 종료되었습니다.');
});
