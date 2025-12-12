// 숫자를 천 단위 구분자로 포맷팅
function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return '0';
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// 기능 플래그: 인증 가드 활성화 (프로덕션 모드)
const ENABLE_AUTH_GUARD = true;

// 테스트 모드: M365 인증 없이 특정 사용자로 가정 (false로 설정하면 실제 로그인 필요)
const TEST_MODE = true;
const TEST_USER_NAME = '김웅희'; // 테스트용 사용자 이름
const ENABLE_SERVER_API = true; // 서버 API 사용 (샘플 데이터 사용 안 함)

// MSAL(Microsoft 365) 설정
const MSAL_CLIENT_ID = 'YOUR_CLIENT_ID_HERE'; // TODO: 실제 앱 등록 Client ID로 교체
const MSAL_TENANT_ID = 'YOUR_TENANT_ID_HERE'; // TODO: 실제 테넌트 ID로 교체 (또는 'common')
const msalConfig = {
    auth: {
        clientId: MSAL_CLIENT_ID,
        authority: `https://login.microsoftonline.com/${MSAL_TENANT_ID}`,
        redirectUri: window.location.origin
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false
    }
};

let msalInstance = null;
let msalAccount = null;
// 현재 로그인한 사용자 정보
let currentUser = null;
let testUserButtonRef = null;

// API 베이스 URL 가져오기 (동적으로 현재 호스트 사용)
function getApiBaseUrl() {
    // file:// 프로토콜로 열린 경우 감지
    if (window.location.protocol === 'file:') {
        console.error('❌ file:// 프로토콜로 열렸습니다. 서버를 통해 접속해야 합니다.');
        return null; // null 반환하여 오류 처리
    }
    
    // 현재 페이지의 호스트와 포트 사용
    const protocol = window.location.protocol;
    const hostname = window.location.hostname;
    const port = window.location.port;
    
    // Live Server (포트 5500) 또는 다른 정적 파일 서버 포트를 사용하는 경우
    // 서버 포트(3000)로 변경
    const STATIC_FILE_SERVER_PORTS = ['5500', '8080', '8000', '5000'];
    const SERVER_PORT = '3000';
    
    // 포트가 정적 파일 서버 포트이면 서버 포트로 변경
    if (port && STATIC_FILE_SERVER_PORTS.includes(port)) {
        const baseUrl = `${protocol}//${hostname}:${SERVER_PORT}`;
        console.log(`📍 정적 파일 서버 포트(${port}) 감지 → 서버 포트(${SERVER_PORT})로 변경`);
        console.log('📍 API 베이스 URL:', baseUrl);
        return baseUrl;
    }
    
    // 포트가 있으면 포함, 없으면 기본 포트 사용 (하지만 서버는 3000 포트)
    // 네트워크 접속 시 포트가 명시되어 있으면 그대로 사용
    if (port && port !== '' && port !== '80' && port !== '443') {
        const baseUrl = `${protocol}//${hostname}:${port}`;
        console.log('📍 API 베이스 URL:', baseUrl);
        return baseUrl;
    }
    
    // 포트가 없으면 기본값으로 3000 사용 (서버 포트)
    const baseUrl = `${protocol}//${hostname}:3000`;
    console.log('📍 API 베이스 URL (기본 포트 3000):', baseUrl);
    return baseUrl;
}

const API_BASE_URL = getApiBaseUrl();

// API_BASE_URL이 null이면 오류 표시
if (!API_BASE_URL) {
    console.error('❌ API 베이스 URL을 가져올 수 없습니다.');
}

// 서버 상태 확인 함수
async function checkServerStatus() {
    if (!ENABLE_SERVER_API) {
        return false;
    }
    
    if (!API_BASE_URL) {
        console.error('❌ API 베이스 URL이 없습니다. 서버를 통해 접속해야 합니다.');
        return false;
    }
    
    try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 2000); // 2초 타임아웃
        
        console.log('🔍 서버 상태 확인:', `${API_BASE_URL}/api/health`);
        const healthCheck = await fetch(`${API_BASE_URL}/api/health`, { 
            method: 'GET',
            signal: controller.signal
        });
        
        clearTimeout(timeoutId);
        
        if (healthCheck.ok) {
            const healthResult = await healthCheck.json();
            if (healthResult.success) {
                console.log('✅ 서버가 실행 중입니다.');
                return true;
            }
        }
        console.warn('⚠️ 서버 헬스체크 실패:', healthCheck.status, healthCheck.statusText);
        return false;
    } catch (error) {
        console.log('⚠️ 서버가 실행 중이지 않습니다:', error.message);
        console.log('📍 현재 URL:', window.location.href);
        console.log('📍 API 베이스 URL:', API_BASE_URL);
        return false;
    }
}

// 서버가 시작될 때까지 대기하는 함수
async function waitForServer(maxWaitTime = 60000, checkInterval = 2000) {
    const startTime = Date.now();
    let attemptCount = 0;
    
    while (Date.now() - startTime < maxWaitTime) {
        attemptCount++;
        const isRunning = await checkServerStatus();
        
        if (isRunning) {
            console.log(`✅ 서버가 준비되었습니다! (${attemptCount}번 시도)`);
            return true;
        }
        
        console.log(`⏳ 서버 대기 중... (${attemptCount}번 시도, ${Math.floor((Date.now() - startTime) / 1000)}초 경과)`);
        await new Promise(resolve => setTimeout(resolve, checkInterval));
    }
    
    console.warn('⚠️ 서버 시작 대기 시간 초과');
    return false;
}

// 서버 시작 안내 함수 (자동 시작 비활성화 - 수동 시작만 안내)
async function tryStartServer() {
    console.log('⚠️ 서버가 실행 중이지 않습니다.');
    console.log('💡 서버를 시작하려면:');
    console.log('   1. 프로젝트 폴더에서 "start-all.cmd" 파일을 더블클릭하세요');
    console.log('   2. 또는 터미널에서 "node server.js" 명령어를 실행하세요');
    console.log('   3. 서버가 시작되면 이 페이지를 새로고침하세요');

    return false;
}

function applyTestUser(name) {
    if (!name) return false;
    const trimmed = String(name).trim();
    if (!trimmed) return false;
    currentUser = {
        name: trimmed,
        username: trimmed,
        displayName: trimmed
    };
    console.log(`🧪 테스트 사용자 설정: ${trimmed}`);
    updateTestUserButtonLabel();
    updateUserDisplay();
    
    // 서버 상태 확인 (안내만 표시, 자동 시작 안 함)
    checkServerStatus().then(serverRunning => {
        if (!serverRunning) {
            console.warn('⚠️ 서버가 실행 중이지 않습니다.');
            console.log('💡 서버를 시작하려면:');
            console.log('   1. 프로젝트 폴더에서 "start-all.cmd" 파일을 더블클릭하세요');
            console.log('   2. 또는 터미널에서 "node server.js" 명령어를 실행하세요');
            console.log('   3. 서버가 실행되면 이 페이지를 새로고침하세요');
        }
    });
    
    return true;
}

function updateTestUserButtonLabel() {
    if (!testUserButtonRef) return;
    if (!TEST_MODE) {
        testUserButtonRef.style.display = 'none';
        return;
    }
    testUserButtonRef.style.display = '';
    const name = currentUser?.name || '미지정';
    testUserButtonRef.textContent = `테스트 사용자: ${name}`;
}

// 사용자 정보 표시 업데이트
function updateUserDisplay() {
    const userDisplay = document.getElementById('current-user-display');
    if (!userDisplay) return;
    
    if (currentUser && currentUser.name) {
        userDisplay.textContent = `사용자: ${currentUser.name}`;
        userDisplay.style.display = '';
    } else {
        userDisplay.style.display = 'none';
    }
}

function showAuthOverlay() {
    const overlay = document.getElementById('auth-overlay');
    if (overlay) overlay.style.display = 'flex';
}

function hideAuthOverlay() {
    const overlay = document.getElementById('auth-overlay');
    if (overlay) overlay.style.display = 'none';
}

function showApp() {
    const appRoot = document.getElementById('app-root');
    if (appRoot) appRoot.style.display = '';
}

function hideApp() {
    const appRoot = document.getElementById('app-root');
    if (appRoot) appRoot.style.display = 'none';
}

async function initializeMsalAndGuard() {
    if (TEST_MODE) {
        applyTestUser(TEST_USER_NAME);
        console.log(`🧪 테스트 모드: ${TEST_USER_NAME} 사용자로 로그인됨`);
        hideAuthOverlay();
        showApp();
        
        // 서버 상태 확인 (비동기로 실행, 사용자 경험 방해하지 않음)
        checkServerStatus().then(serverRunning => {
            if (!serverRunning) {
                console.warn('⚠️ 서버가 실행 중이지 않습니다. 데이터 조회 전에 서버를 실행해주세요.');
            }
        });
        
        return;
    }
    
    if (!ENABLE_AUTH_GUARD) {
        // 가드 비활성화: 오버레이 숨기고 앱 즉시 표시
        hideAuthOverlay();
        showApp();
        return;
    }
    if (!window.msal) {
        console.error('MSAL 라이브러리가 로드되지 않았습니다.');
        // MSAL 설정이 올바르지 않은 경우 테스트 모드로 전환
        if (MSAL_CLIENT_ID === 'YOUR_CLIENT_ID_HERE' || MSAL_TENANT_ID === 'YOUR_TENANT_ID_HERE') {
            console.warn('MSAL 설정이 완료되지 않았습니다. 테스트 모드로 전환합니다.');
            if (TEST_MODE) {
                applyTestUser(TEST_USER_NAME);
                hideAuthOverlay();
                showApp();
                return;
            }
        }
        return;
    }

    // MSAL 설정 확인
    if (MSAL_CLIENT_ID === 'YOUR_CLIENT_ID_HERE' || MSAL_TENANT_ID === 'YOUR_TENANT_ID_HERE') {
        console.warn('MSAL 설정이 완료되지 않았습니다. 테스트 모드로 전환합니다.');
        if (TEST_MODE) {
            applyTestUser(TEST_USER_NAME);
            hideAuthOverlay();
            showApp();
            return;
        } else {
            alert('Microsoft 365 로그인 설정이 완료되지 않았습니다.\n\n관리자에게 문의하거나, 테스트 모드를 활성화해주세요.');
            return;
        }
    }

    msalInstance = new msal.PublicClientApplication(msalConfig);

    try {
        const redirectResult = await msalInstance.handleRedirectPromise();
        if (redirectResult && redirectResult.account) {
            msalInstance.setActiveAccount(redirectResult.account);
        }
    } catch (e) {
        console.error('MSAL redirect 처리 오류:', e);
    }

    const accounts = msalInstance.getAllAccounts();
    msalAccount = accounts && accounts.length > 0 ? accounts[0] : null;

    if (msalAccount) {
        msalInstance.setActiveAccount(msalAccount);
        // M365 계정에서 사용자 정보 추출
        currentUser = {
            name: msalAccount.name || msalAccount.username,
            username: msalAccount.username,
            displayName: msalAccount.name || msalAccount.username
        };
        updateTestUserButtonLabel();
        updateUserDisplay();
        hideAuthOverlay();
        showApp();
    } else {
        hideApp();
        showAuthOverlay();
    }
}

async function loginWithM365() {
    if (!msalInstance) {
        // MSAL 설정이 올바르지 않은 경우
        if (MSAL_CLIENT_ID === 'YOUR_CLIENT_ID_HERE' || MSAL_TENANT_ID === 'YOUR_TENANT_ID_HERE') {
            alert('Microsoft 365 로그인 설정이 완료되지 않았습니다.\n\n관리자에게 문의하거나, 테스트 모드를 사용해주세요.');
            console.error('MSAL 설정이 완료되지 않았습니다. script.js 파일에서 MSAL_CLIENT_ID와 MSAL_TENANT_ID를 설정해주세요.');
            return;
        }
        alert('로그인 시스템을 초기화할 수 없습니다. 페이지를 새로고침해주세요.');
        return;
    }
    try {
        const result = await msalInstance.loginPopup({ scopes: ['User.Read'] });
        if (result && result.account) {
            msalInstance.setActiveAccount(result.account);
            msalAccount = result.account;
            // M365 계정에서 사용자 정보 추출
            currentUser = {
                name: result.account.name || result.account.username,
                username: result.account.username,
                displayName: result.account.name || result.account.username
            };
            updateTestUserButtonLabel();
            updateUserDisplay();
            hideAuthOverlay();
            showApp();
        }
    } catch (e) {
        alert('로그인에 실패했습니다. 다시 시도해주세요.\n\n오류: ' + (e.message || '알 수 없는 오류'));
        console.error('MSAL loginPopup 오류:', e);
    }
}

// 현재 사용자 정보 가져오기
function getCurrentUser() {
    return currentUser;
}

// 로그아웃 함수
function logout() {
    // 테스트 모드인 경우
    if (TEST_MODE) {
        currentUser = null;
        updateUserDisplay();
        updateTestUserButtonLabel();
        hideApp();
        showAuthOverlay();
        alert('로그아웃되었습니다.');
        return;
    }
    
    // M365 로그아웃
    if (msalInstance) {
        try {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                msalInstance.logoutPopup({
                    account: accounts[0]
                });
            }
        } catch (e) {
            console.error('로그아웃 오류:', e);
        }
    }
    
    // 사용자 정보 초기화
    currentUser = null;
    msalAccount = null;
    updateUserDisplay();
    updateTestUserButtonLabel();
    hideApp();
    showAuthOverlay();
}

// DOM 요소들 가져오기
const tabItems = document.querySelectorAll('.tab-item');
const queryBtn = document.getElementById('query-btn');
const m365LoginBtn = document.getElementById('m365-login-btn');
const testUserBtn = document.getElementById('test-user-btn');
const logoutBtn = document.getElementById('logout-btn');
const monthlySummaryDownloadBtn = document.getElementById('monthly-summary-download-btn');
const settledDownloadBtn = document.getElementById('settled-download-btn');
const unsettledDownloadBtn = document.getElementById('unsettled-download-btn');
const periodInput = document.getElementById('period');

if (testUserBtn) {
    testUserButtonRef = testUserBtn;
    if (!TEST_MODE) {
        testUserBtn.style.display = 'none';
    } else {
        updateTestUserButtonLabel();
        testUserBtn.addEventListener('click', () => {
            const defaultName = currentUser?.name || TEST_USER_NAME;
            const input = prompt('테스트 사용자 이름을 입력하세요.', defaultName);
            if (!input) {
                return;
            }
            const success = applyTestUser(input);
            if (success) {
                alert(`테스트 사용자를 "${currentUser.name}"(으)로 변경했습니다. 조회 버튼을 눌러 데이터를 새로 불러오세요.`);
            } else {
                alert('올바른 사용자 이름을 입력해주세요.');
            }
        });
    }
}

// 로그아웃 버튼 이벤트 리스너
if (logoutBtn) {
    logoutBtn.addEventListener('click', () => {
        if (confirm('로그아웃하시겠습니까?')) {
            logout();
        }
    });
}

let latestServerData = {
    settled: {
        monthly: [],
        detail: []
    },
    unsettled: {
        amount: 0,
        detail: []
    }
};

let currentFilteredMonthlyData = [];
let currentFilteredSettledDetail = [];
let currentFilteredUnsettledDetail = [];

// 원본 데이터 저장 (필터링 전)
let originalMonthlyData = [];
let originalSettledDetail = [];
let originalUnsettledDetail = [];

// 정렬 상태 추적
let sortState = {
    monthly: { column: null, direction: null },
    settled: { column: null, direction: null },
    unsettled: { column: null, direction: null }
};




// 조회 기간에 따른 데이터 필터링 함수들
function parsePeriod(periodStr) {
    const match = periodStr.match(/(\d{4})-(\d{2})\s*~\s*(\d{4})-(\d{2})/);
    if (!match) return null;
    
    const [, startYear, startMonth, endYear, endMonth] = match;
    return {
        start: `${startYear}-${startMonth}`,
        end: `${endYear}-${endMonth}`,
        startYear: parseInt(startYear),
        startMonth: parseInt(startMonth),
        endYear: parseInt(endYear),
        endMonth: parseInt(endMonth)
    };
}

function isMonthInRange(monthStr, period) {
    if (!period) return true;
    
    const [year, month] = monthStr.split('-').map(Number);
    const monthNum = year * 12 + month;
    const startNum = period.startYear * 12 + period.startMonth;
    const endNum = period.endYear * 12 + period.endMonth;
    
    return monthNum >= startNum && monthNum <= endNum;
}

function filterDataByPeriod(data, period) {
    if (!period) return data;
    
    // period 파싱: 문자열이면 파싱, 객체면 그대로 사용
    const parsedPeriod = typeof period === 'string' ? parsePeriod(period) : period;
    if (!parsedPeriod) return data;
    
    // 🔥 정산월(month) 기준으로 필터링
    // 조회기간 2025-01~2025-02 → 정산월이 2025-01~2025-02인 데이터
    console.log(`📅 필터링 범위 계산:`, {
        조회기간: `${parsedPeriod.startYear}-${String(parsedPeriod.startMonth).padStart(2, '0')} ~ ${parsedPeriod.endYear}-${String(parsedPeriod.endMonth).padStart(2, '0')}`,
        정산월범위: `${parsedPeriod.startYear}-${String(parsedPeriod.startMonth).padStart(2, '0')} ~ ${parsedPeriod.endYear}-${String(parsedPeriod.endMonth).padStart(2, '0')}`
    });
    
    return data.filter(item => {
        // 🔥 정산 상세 내역 데이터: 정산월(month) 기준으로 필터링
        // 조회기간 2025-01~2025-02 → 정산월이 2025-01~2025-02인 데이터
        const settlementMonth = item.month || item.settlementMonth || '';
        
        if (!settlementMonth) {
            return false;
        }
        
        // 정산월이 조회기간 범위에 있는지 확인
        const isInRange = isMonthInRange(settlementMonth, parsedPeriod);
        
        // 🔍 디버깅: 정산월 필터링 확인 (처음 20개 또는 2024-12 데이터)
        const itemIndex = data.indexOf(item);
        if (itemIndex < 20 || settlementMonth === '2024-12') {
            console.log(`   📋 항목 ${itemIndex + 1}: 정산월=${settlementMonth}, 포함=${isInRange}, 조회기간=${parsedPeriod.startYear}-${String(parsedPeriod.startMonth).padStart(2, '0')}~${parsedPeriod.endYear}-${String(parsedPeriod.endMonth).padStart(2, '0')}`);
            if (settlementMonth === '2024-12' && !isInRange) {
                console.error(`   ❌ [오류] 2024-12 데이터가 필터링에서 제외됨!`);
            }
        }
        
        return isInRange;
    });
}

function calculateUnsettledAmount(detailData) {
    return detailData.reduce((sum, item) => sum + item.amount, 0);
}

// 데이터 정렬
function sortTableData(data, column, direction) {
    if (!column || !direction) {
        return data;
    }
    
    const sorted = [...data].sort((a, b) => {
        let aVal = a[column];
        let bVal = b[column];
        
        // 숫자 필드 처리
        if (column === 'amount') {
            aVal = Number(aVal) || 0;
            bVal = Number(bVal) || 0;
            return direction === 'asc' ? aVal - bVal : bVal - aVal;
        }
        
        // 날짜 필드 처리 (YYYY-MM-DD 형식)
        if (column === 'paymentDate') {
            // YYYY-MM-DD 형식은 문자열 비교로도 올바르게 정렬됨
            aVal = String(aVal || '');
            bVal = String(bVal || '');
            if (direction === 'asc') {
                return aVal.localeCompare(bVal);
            } else {
                return bVal.localeCompare(aVal);
            }
        }
        
        // 문자열 필드 처리
        aVal = String(aVal || '').toLowerCase();
        bVal = String(bVal || '').toLowerCase();
        
        if (direction === 'asc') {
            return aVal.localeCompare(bVal);
        } else {
            return bVal.localeCompare(aVal);
        }
    });
    
    return sorted;
}

// 숫자를 천 단위 구분자로 포맷팅 (예: 1700000 -> "1,700,000")
function formatNumber(num) {
  if (num === null || num === undefined || isNaN(num)) return '0';
  return num.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}

// 통화 문자열을 숫자로 변환 (예: "1,700,000원" -> 1700000)
function parseCurrencyToNumber(text) {
    if (typeof text !== 'string') return null;
    const cleaned = text
        .replace(/\s+/g, '')
        .replace(/[,]/g, '')
        .replace(/원$/,'');
    if (cleaned === '') return null;
    const value = Number(cleaned);
    return Number.isFinite(value) ? value : null;
}

// 테이블 데이터를 엑셀로 다운로드
function downloadTableAsExcel(tableId, filename) {
    const table = document.getElementById(tableId);
    if (!table) {
        alert('테이블을 찾을 수 없습니다.');
        return;
    }

    // 테이블 데이터 추출
    const data = [];
    const rows = table.querySelectorAll('tr');
    
    rows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('td, th');
        cells.forEach(cell => {
            const text = cell.textContent.trim();
            // 금액 형식이면 숫자로 변환하여 삽입
            const num = parseCurrencyToNumber(text);
            rowData.push(num !== null ? num : text);
        });
        if (rowData.length > 0) {
            data.push(rowData);
        }
    });

    if (data.length === 0) {
        alert('다운로드할 데이터가 없습니다.');
        return;
    }

    // 워크북 생성
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);

    // 컬럼 너비 설정
    const colWidths = [];
    data[0].forEach((_, index) => {
        colWidths.push({ wch: 15 });
    });
    ws['!cols'] = colWidths;

    // 워크시트를 워크북에 추가
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // 파일 다운로드
    XLSX.writeFile(wb, filename);
}


// 선택된 월 상태
let selectedMonth = null;

// 상세 내역 데이터에서 정산월 기준 월별 집계 계산 (공통 함수)
// 🔥 N열 정산월(month) 필드 기준으로만 집계
function calculateMonthlySummaryFromDetail(detailData) {
    const monthlyMap = new Map();
    
    if (!Array.isArray(detailData) || detailData.length === 0) {
        console.log('⚠️ calculateMonthlySummaryFromDetail: 데이터 없음');
        return [];
    }
    
    let monthNullCount = 0;
    let monthValidCount = 0;
    
    detailData.forEach((item, idx) => {
        // 🔥 정산월은 item.month 필드를 우선 사용 (N열에서 읽은 값)
        const month = item.month || item.settlementMonth || null;
        
        if (!month) {
            monthNullCount++;
            // 디버깅: 처음 5개만 로그
            if (idx < 5) {
                console.warn(`   ⚠️ [프론트엔드 월별집계] index=${idx}: 정산월 없음, item.month="${item.month}", item.settlementMonth="${item.settlementMonth}", amount=${item.amount}`);
            }
            return;
        }
        
        // 빈 문자열 체크
        if (String(month).trim() === '') {
            monthNullCount++;
            if (idx < 5) {
                console.warn(`   ⚠️ [프론트엔드 월별집계] index=${idx}: 정산월 빈 문자열, amount=${item.amount}`);
            }
            return;
        }
        
        // 미정산 데이터 제외
        if (month.includes('미정산') || month.includes('_미정산')) {
            return;
        }
        
        monthValidCount++;
        const amount = typeof item.amount === 'number' 
            ? item.amount 
            : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
        
        if (monthlyMap.has(month)) {
            monthlyMap.set(month, monthlyMap.get(month) + amount);
        } else {
            monthlyMap.set(month, amount);
        }
        
        // 디버깅: 처음 10개만 로그
        if (idx < 10) {
            console.log(`   [프론트엔드 월별집계] index=${idx}: 정산월="${month}", 금액=${amount}, 누적합계=${monthlyMap.get(month)}`);
        }
    });
    
    console.log(`\n📊 프론트엔드 월별 집계 통계:`);
    console.log(`   ✅ 정산월 있음: ${monthValidCount}개`);
    console.log(`   ⚠️ 정산월 없음: ${monthNullCount}개`);
    console.log(`   📋 월별 집계 결과: ${monthlyMap.size}개 월`);
    monthlyMap.forEach((amount, month) => {
        console.log(`      ${month}: ${amount.toLocaleString()}원`);
    });
    
    // Map을 배열로 변환하고 정렬 (정산월 내림차순: 가장 최근월부터)
    return Array.from(monthlyMap.entries())
        .map(([month, amount]) => ({ month, amount }))
        .sort((a, b) => (b.month || '').localeCompare(a.month || ''));
}

// 월별 정산 요약 테이블 업데이트
// 🔥 항상 현재 상세 내역 데이터에서 정산월 기준으로 계산하여 표시
function updateMonthlySummary(data = null) {
    // data가 제공되지 않으면 현재 상세 내역에서 계산
    let monthlyData;
    if (data === null) {
        // 현재 필터링된 상세 내역 데이터 가져오기
        const detailData = currentFilteredSettledDetail.length > 0
            ? currentFilteredSettledDetail
            : (originalSettledDetail.length > 0 ? originalSettledDetail : latestServerData.settled?.detail || []);
        
        // 상세 내역에서 정산월 기준으로 집계 계산
        monthlyData = calculateMonthlySummaryFromDetail(detailData);
    } else if (Array.isArray(data) && data.length === 0) {
        // 빈 배열이 명시적으로 전달된 경우 (예: 미정산 탭 선택 시)
        monthlyData = [];
    } else {
        // 외부에서 계산된 데이터 사용 (하지만 항상 상세 내역에서 계산된 데이터여야 함)
        monthlyData = data;
    }
    
    // 원본 데이터 저장 (정렬을 위해)
    originalMonthlyData = [...monthlyData];
    
    let displayData = [...monthlyData];
    if (sortState.monthly.column) {
        displayData = sortTableData(displayData, sortState.monthly.column, sortState.monthly.direction);
    } else {
        // 정렬 상태가 없을 때 기본적으로 정산월 내림차순 정렬 (가장 최근월부터)
        displayData = sortTableData(displayData, 'month', 'desc');
    }
    
    const tbody = document.getElementById('monthly-summary-tbody');
    tbody.innerHTML = '';
    
    if (displayData.length === 0) {
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="2" class="table-placeholder">조회 결과가 없습니다.</td>';
        tbody.appendChild(row);
        document.getElementById('total-settled').textContent = '0';
        setupResizeHandlesAfterUpdate();
        return;
    }
    
    // 가장 긴 정산금액 문자열 길이 계산
    let maxAmountLength = 0;
    let total = 0;
    displayData.forEach(item => {
        const formattedAmount = formatNumber(item.amount);
        if (formattedAmount.length > maxAmountLength) {
            maxAmountLength = formattedAmount.length;
        }
        total += item.amount;
    });
    const formattedTotal = formatNumber(total);
    if (formattedTotal.length > maxAmountLength) {
        maxAmountLength = formattedTotal.length;
    }
    
    // 정산금액 열 동적 너비 설정
    setDynamicColumnWidth('#monthly-summary-table', 2, maxAmountLength, 'monthly-summary-amount-column-style', true);
    
    displayData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item.month}</td>
            <td>${formatNumber(item.amount)}</td>
        `;
        // 행 클릭 시 해당 월로 상세내역 필터
        row.style.cursor = 'pointer';
        row.addEventListener('click', () => {
            // 선택 행 하이라이트 처리
            const allRows = tbody.querySelectorAll('tr');
            allRows.forEach(r => r.style.backgroundColor = '');
            row.style.backgroundColor = '#fffbe6';

            selectedMonth = item.month;
            // 현재 조회 기간이 있으면 필터링된 데이터에서 해당 월 찾기
            const baseDetailData = currentFilteredSettledDetail.length > 0
                ? currentFilteredSettledDetail
                : latestServerData.settled.detail || [];
            // 백엔드에서 정산월 컬럼 값을 month 필드로 보내주므로 그대로 사용
            const filtered = baseDetailData.filter(detail => {
                return detail.month === selectedMonth;
            });
            updateSettledDetail(filtered);
        });

        // 초기 선택 유지 시 하이라이트
        if (selectedMonth && selectedMonth === item.month) {
            row.style.backgroundColor = '#fffbe6';
        }

        tbody.appendChild(row);
    });
    
    document.getElementById('total-settled').textContent = formatNumber(total);
    
    // 정렬 헤더 UI 업데이트
    updateSortHeaders('monthly');
    
    // 리사이즈 핸들 재설정
    setupResizeHandlesAfterUpdate();
    
    // 합계 일관성 검증 (상세 내역이 이미 업데이트된 경우)
    setTimeout(() => {
        validateSettlementTotals();
    }, 100);
}

// 월별 정산 요약과 상세 내역 합계 일관성 검증
function validateSettlementTotals() {
    try {
        // 1. 월별 정산 요약 테이블의 합계 계산 (정산월 기준 합계)
        let monthlySummaryTotal = 0;
        if (Array.isArray(originalMonthlyData) && originalMonthlyData.length > 0) {
            monthlySummaryTotal = originalMonthlyData.reduce((sum, item) => {
                const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
                return sum + amount;
            }, 0);
        }
        
        // 2. 월 정산 상세 내역에서 정산월 기준 합계 계산
        const detailData = currentFilteredSettledDetail.length > 0
            ? currentFilteredSettledDetail
            : (originalSettledDetail.length > 0 ? originalSettledDetail : latestServerData.settled?.detail || []);
        
        // 상세 내역을 정산월 기준으로 집계
        const detailMonthlyMap = new Map();
        if (Array.isArray(detailData) && detailData.length > 0) {
            detailData.forEach(item => {
                const month = item.month || item.settlementMonth || null;
                if (!month) return;
                // 미정산 데이터 제외
                if (month.includes('미정산') || month.includes('_미정산')) {
                    return;
                }
                const amount = typeof item.amount === 'number' ? item.amount : parseFloat(String(item.amount || 0).replace(/[^0-9.-]/g, '')) || 0;
                if (detailMonthlyMap.has(month)) {
                    detailMonthlyMap.set(month, detailMonthlyMap.get(month) + amount);
                } else {
                    detailMonthlyMap.set(month, amount);
                }
            });
        }
        
        // 정산월 기준 집계 합계 계산
        const detailTotal = Array.from(detailMonthlyMap.values()).reduce((sum, amount) => sum + amount, 0);
        
        // 3. 두 합계 비교 (소수점 오차 허용: 0.01원 이내)
        const difference = Math.abs(monthlySummaryTotal - detailTotal);
        const isMatch = difference < 0.01;
        
        if (isMatch) {
            console.log(`✅ 합계 일관성 검증 통과: 월별 정산 요약 합계(${monthlySummaryTotal.toLocaleString()}원) = 상세 내역 합계(${detailTotal.toLocaleString()}원)`);
        } else {
            console.warn(`⚠️ 합계 일관성 검증 실패!`);
            console.warn(`   월별 정산 요약 합계: ${monthlySummaryTotal.toLocaleString()}원 (${originalMonthlyData.length}개 월)`);
            console.warn(`   월 정산 상세 내역 합계: ${detailTotal.toLocaleString()}원 (${detailData.length}개 항목, ${detailMonthlyMap.size}개 월)`);
            console.warn(`   차이: ${difference.toLocaleString()}원`);
            
            // 월별 상세 비교
            console.warn(`   월별 상세 비교:`);
            const monthlySummaryMap = new Map(originalMonthlyData.map(item => [item.month, item.amount]));
            detailMonthlyMap.forEach((amount, month) => {
                const summaryAmount = monthlySummaryMap.get(month) || 0;
                const monthDiff = Math.abs(summaryAmount - amount);
                if (monthDiff >= 0.01) {
                    console.warn(`     ${month}: 요약=${summaryAmount.toLocaleString()}원, 상세=${amount.toLocaleString()}원, 차이=${monthDiff.toLocaleString()}원`);
                }
            });
        }
        
        return isMatch;
    } catch (error) {
        console.error('❌ 합계 일관성 검증 중 오류:', error);
        return false;
    }
}

// 미정산 금액 합계 카드 업데이트
function updateUnsettledSummaryTable(data, totalAmountOverride = null) {
    const totalEl = document.getElementById('unsettled-total-value');
    if (!totalEl) {
        console.error('❌ unsettled-total-value 요소를 찾을 수 없습니다.');
        return;
    }

    const totalAmount = totalAmountOverride !== null
        ? totalAmountOverride
        : (data || []).reduce((sum, item) => sum + (item.amount || 0), 0);

    totalEl.textContent = formatNumber(totalAmount);
}

// 정산 상세 내역 테이블 업데이트
function updateSettledDetail(data, skipOriginalSave = false) {
    // 원본 데이터 저장 (필터 적용 전 데이터)
    if (!skipOriginalSave) {
        originalSettledDetail = [...(data || [])];
    }
    
    let displayData = [...(data || [])];
    
    // 필터 적용
    displayData = applyFiltersToData(displayData);
    
    if (sortState.settled.column) {
        displayData = sortTableData(displayData, sortState.settled.column, sortState.settled.direction);
    }
    
    const tbody = document.getElementById('settled-detail-tbody');
    tbody.innerHTML = '';
    
    if (!displayData || displayData.length === 0) {
        console.log('⚠️ updateSettledDetail: 데이터가 없습니다.');
        const row = document.createElement('tr');
        row.innerHTML = '<td colspan="6" class="table-placeholder">조회 결과가 없습니다.</td>';
        tbody.appendChild(row);
        document.getElementById('totalAmountCell').textContent = '0';
        setupResizeHandlesAfterUpdate();
        return;
    }
    
    console.log('📋 updateSettledDetail 호출됨, 데이터 개수:', displayData.length);
    console.log('📋 첫 번째 항목:', displayData[0]);
    
    // 각 컬럼별 최대 문자열 길이 계산
    let maxMonthLength = 0;
    let maxPaymentDateLength = 0;
    let maxMerchantLength = 0;
    let maxAccountNameLength = 0;
    let maxAmountLength = 0;
    let maxNoteLength = 0;
    let total = 0;
    
    displayData.forEach(item => {
        // 정산월
        const monthValue = String(item.settlementMonth || item.month || '');
        if (monthValue.length > maxMonthLength) {
            maxMonthLength = monthValue.length;
        }
        
        // 지급일
        const paymentDateValue = String(item.paymentDate || '');
        if (paymentDateValue.length > maxPaymentDateLength) {
            maxPaymentDateLength = paymentDateValue.length;
        }
        
        // 사용처
        const merchantValue = String(item.merchant || '');
        if (merchantValue.length > maxMerchantLength) {
            maxMerchantLength = merchantValue.length;
        }
        
        // 계정명
        const accountNameValue = String(item.accountName || '');
        if (accountNameValue.length > maxAccountNameLength) {
            maxAccountNameLength = accountNameValue.length;
        }
        
        // 정산금액
        const formattedAmount = formatNumber(item.amount);
        if (formattedAmount.length > maxAmountLength) {
            maxAmountLength = formattedAmount.length;
        }
        
        // 비고
        const noteText = item.note ? String(item.note) : '';
        if (noteText.length > maxNoteLength) {
            maxNoteLength = noteText.length;
        }
        
        total += item.amount;
    });
    
    // 합계 금액 길이도 고려
    const formattedTotal = formatNumber(total);
    if (formattedTotal.length > maxAmountLength) {
        maxAmountLength = formattedTotal.length;
    }
    
    // 각 컬럼 너비 설정 (비고 제외)
    setDynamicColumnWidth('#settled-detail-table', 1, maxMonthLength, 'settled-month-column-style', false, 80);
    setDynamicColumnWidth('#settled-detail-table', 2, maxPaymentDateLength, 'settled-paymentDate-column-style', false, 100);
    setDynamicColumnWidth('#settled-detail-table', 3, maxMerchantLength, 'settled-merchant-column-style', false, 120);
    setDynamicColumnWidth('#settled-detail-table', 4, maxAccountNameLength, 'settled-accountName-column-style', false, 120);
    setDynamicColumnWidth('#settled-detail-table', 5, maxAmountLength, 'settled-amount-column-style', true, 100);
    setDynamicColumnWidth('#settled-detail-table', 6, maxNoteLength, 'settled-note-column-style', false, 120);
    
    console.log('📊 컬럼 너비 설정:', {
        정산월: maxMonthLength,
        지급일: maxPaymentDateLength,
        사용처: maxMerchantLength,
        계정명: maxAccountNameLength,
        정산금액: maxAmountLength,
        비고: maxNoteLength
    });
    
    let totalAmount = 0;
    displayData.forEach((item, index) => {
        // 정산월 값 확인
        const monthValue = item.settlementMonth || item.month || '';
        if (index === 0) {
            console.log(`🔍 첫 번째 행의 정산월 값: "${monthValue}"`);
        }
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${monthValue}</td>
            <td>${item.paymentDate || ''}</td>
            <td>${item.merchant || ''}</td>
            <td>${item.accountName || ''}</td>
            <td>${formatNumber(item.amount)}</td>
            <td>${item.note || ''}</td>
        `;
        tbody.appendChild(row);
        totalAmount += item.amount;
    });
    
    // 테이블 높이를 7개 행(280px)으로 고정하기 위해 빈 행 추가
    const targetRowCount = 7;
    const currentRowCount = displayData.length;
    if (currentRowCount < targetRowCount) {
        for (let i = currentRowCount; i < targetRowCount; i++) {
            const emptyRow = document.createElement('tr');
            emptyRow.innerHTML = `
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
                <td></td>
            `;
            emptyRow.style.height = '40px';
            emptyRow.style.minHeight = '40px';
            emptyRow.style.maxHeight = '40px';
            tbody.appendChild(emptyRow);
        }
    }
    
    document.getElementById('totalAmountCell').textContent = formatNumber(totalAmount);
    console.log('✅ updateSettledDetail 완료, 총합:', totalAmount);
    
    // 정렬 헤더 UI 업데이트
    updateSortHeaders('settled');
    
    // 필터 아이콘 상태 업데이트
    const settledTable = document.getElementById('settled-detail-table');
    if (settledTable) {
        const filterIcons = settledTable.querySelectorAll('.filter-icon');
        filterIcons.forEach(icon => {
            const column = icon.getAttribute('data-column');
            updateFilterIconState(icon, column);
        });
    }
    
    // 리사이즈 핸들 재설정
    setupResizeHandlesAfterUpdate();
    
    // 합계 일관성 검증 (월별 정산 요약이 이미 업데이트된 경우)
    setTimeout(() => {
        validateSettlementTotals();
    }, 100);
}

// 미정산 상세 내역 테이블 업데이트
function updateUnsettledDetail(data) {
    // 원본 데이터 저장
    originalUnsettledDetail = [...(data || [])];
    
    let displayData = [...(data || [])];
    if (sortState.unsettled.column) {
        displayData = sortTableData(displayData, sortState.unsettled.column, sortState.unsettled.direction);
    }
    
    console.log('📋 updateUnsettledDetail 호출됨, 데이터:', displayData);
    const tbody = document.getElementById('unsettled-detail-tbody');
    if (!tbody) {
        console.error('❌ unsettled-detail-tbody 요소를 찾을 수 없습니다.');
        return;
    }
    tbody.innerHTML = '';
    
    if (!displayData || displayData.length === 0) {
        console.log('⚠️ updateUnsettledDetail: 데이터가 없습니다.');
        // 데이터가 없을 때 합계를 0으로 설정
        document.getElementById('total-unsettled-detail').textContent = formatNumber(0);
        const unsettledTotalValue = document.getElementById('unsettled-total-value');
        if (unsettledTotalValue) {
            unsettledTotalValue.textContent = formatNumber(0);
        }
        // 빈 상태 메시지 표시
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = '<td colspan="6" class="table-placeholder">조회 결과가 없습니다.</td>';
        tbody.appendChild(emptyRow);
        setupResizeHandlesAfterUpdate();
        return;
    }
    
    console.log('✅ updateUnsettledDetail: 데이터 처리 시작,', displayData.length, '개 항목');
    
    // 각 컬럼별 최대 문자열 길이 계산
    let maxMonthLength = 0;
    let maxPaymentDateLength = 0;
    let maxMerchantLength = 0;
    let maxAccountNameLength = 0;
    let maxAmountLength = 0;
    let maxNoteLength = 0;
    let total = 0;
    
    displayData.forEach(item => {
        // 정산월
        const monthValue = String(item.settlementMonth || item.month || '');
        if (monthValue.length > maxMonthLength) {
            maxMonthLength = monthValue.length;
        }
        
        // 지급예정일
        const paymentDateValue = String(item.paymentDate || '');
        if (paymentDateValue.length > maxPaymentDateLength) {
            maxPaymentDateLength = paymentDateValue.length;
        }
        
        // 사용처
        const merchantValue = String(item.merchant || '');
        if (merchantValue.length > maxMerchantLength) {
            maxMerchantLength = merchantValue.length;
        }
        
        // 계정명
        const accountNameValue = String(item.accountName || '');
        if (accountNameValue.length > maxAccountNameLength) {
            maxAccountNameLength = accountNameValue.length;
        }
        
        // 정산금액
        const formattedAmount = formatNumber(item.amount);
        if (formattedAmount.length > maxAmountLength) {
            maxAmountLength = formattedAmount.length;
        }
        
        // 비고
        const noteText = item.note ? String(item.note) : '';
        if (noteText.length > maxNoteLength) {
            maxNoteLength = noteText.length;
        }
        
        total += item.amount;
    });
    
    // 합계 금액 길이도 고려
    const formattedTotal = formatNumber(total);
    if (formattedTotal.length > maxAmountLength) {
        maxAmountLength = formattedTotal.length;
    }
    
    // 각 컬럼 너비 설정 (비고 제외)
    setDynamicColumnWidth('#unsettled-detail-table', 1, maxMonthLength, 'unsettled-month-column-style', false, 80);
    setDynamicColumnWidth('#unsettled-detail-table', 2, maxPaymentDateLength, 'unsettled-paymentDate-column-style', false, 100);
    setDynamicColumnWidth('#unsettled-detail-table', 3, maxMerchantLength, 'unsettled-merchant-column-style', false, 120);
    setDynamicColumnWidth('#unsettled-detail-table', 4, maxAccountNameLength, 'unsettled-accountName-column-style', false, 120);
    setDynamicColumnWidth('#unsettled-detail-table', 5, maxAmountLength, 'unsettled-amount-column-style', true, 100);
    setDynamicColumnWidth('#unsettled-detail-table', 6, maxNoteLength, 'unsettled-note-column-style', false, 120);
    
    console.log('📊 미정산 컬럼 너비 설정:', {
        정산월: maxMonthLength,
        지급예정일: maxPaymentDateLength,
        사용처: maxMerchantLength,
        계정명: maxAccountNameLength,
        정산금액: maxAmountLength,
        비고: maxNoteLength
    });
    
    let totalAmount = 0;
    displayData.forEach(item => {
        const monthValue = item.settlementMonth || item.month || '';
        const paymentDate = item.paymentDate || item.date || ''; // 지급예정일은 현재 데이터에 없으므로 빈 값
        const merchant = item.merchant || '';
        const accountName = item.accountName || '';
        const note = item.note || '';
        const amount = item.amount || 0;
        
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${monthValue}</td>
            <td>${paymentDate}</td>
            <td>${merchant}</td>
            <td>${accountName}</td>
            <td>${formatNumber(amount)}</td>
            <td>${note}</td>
        `;
        tbody.appendChild(row);
        totalAmount += amount;
    });
    
    document.getElementById('total-unsettled-detail').textContent = formatNumber(totalAmount);
    const unsettledTotalValue = document.getElementById('unsettled-total-value');
    if (unsettledTotalValue) {
        unsettledTotalValue.textContent = formatNumber(totalAmount);
    }
    
    // 정렬 헤더 UI 업데이트
    updateSortHeaders('unsettled');
    
    // 리사이즈 핸들 재설정
    setupResizeHandlesAfterUpdate();
}

function setDynamicColumnWidth(tableSelector, columnIndex, maxLength, styleId, alignRight = false, minWidth = 80) {
    const tableExists = document.querySelector(tableSelector);
    if (!tableExists) return;
    
    const width = Math.max(maxLength * 8 + 20, minWidth);
    
    // 합계 셀 ID 확인
    let totalCellSelector = '';
    if (columnIndex === 5) {
        // 상세 내역 테이블의 정산금액 열 (5번째)
        if (tableSelector === '#settled-detail-table') {
            totalCellSelector = '#totalAmountCell';
        } else if (tableSelector === '#unsettled-detail-table') {
            totalCellSelector = '#total-unsettled-detail';
        }
    } else if (columnIndex === 2 && tableSelector === '#monthly-summary-table') {
        // 월별 정산 요약 테이블의 정산금액 열 (2번째)
        totalCellSelector = '#total-settled';
    }
    
    const style = document.createElement('style');
    let styleContent = `
        ${tableSelector} th:nth-child(${columnIndex}),
        ${tableSelector} tbody td:nth-child(${columnIndex}),
        ${tableSelector} tfoot td:nth-child(${columnIndex})`;
    
    if (totalCellSelector) {
        styleContent += `,
        ${totalCellSelector}`;
    }
    
    styleContent += ` {
            width: ${width}px !important;
            min-width: ${width}px;
            ${alignRight ? 'text-align: right;' : 'text-align: left;'}
        }`;
    
    style.textContent = styleContent;
    const existingStyle = document.getElementById(styleId);
    if (existingStyle) {
        existingStyle.remove();
    }
    style.id = styleId;
    document.head.appendChild(style);
}

// 탭 변경 이벤트
tabItems.forEach(item => {
    item.addEventListener('click', () => {
        // 모든 탭에서 active 클래스 제거
        tabItems.forEach(tab => tab.classList.remove('active'));
        // 클릭된 탭에 active 클래스 추가
        item.classList.add('active');
        
        // 라디오 버튼 체크
        const radio = item.querySelector('input[type="radio"]');
        radio.checked = true;
        
        // 탭에 따른 데이터 표시
        const tabValue = radio.value;
        if (tabValue === 'settled') {
            // 🔥 월별 정산 요약: 항상 현재 상세 내역에서 계산 (함수 내부에서 처리)
            updateMonthlySummary();
            // 선택된 월이 있으면 필터 적용, 없으면 전체
            // 백엔드에서 정산월 컬럼 값을 month 필드로 보내주므로 그대로 사용
            const finalDetailData = selectedMonth ? 
                detailData.filter(item => {
                    return item.month === selectedMonth;
                }) : 
                detailData;
            updateSettledDetail(finalDetailData);
        } else if (tabValue === 'unsettled') {
            updateMonthlySummary([]);
            updateSettledDetail([]);
        }
    });
});

// 조회 버튼 클릭 이벤트
queryBtn.addEventListener('click', async () => {
    // 드롭다운에서 값을 읽어서 period 문자열 생성
    const startYear = document.getElementById('start-year')?.value || '';
    const startMonth = document.getElementById('start-month')?.value || '';
    const endYear = document.getElementById('end-year')?.value || '';
    const endMonth = document.getElementById('end-month')?.value || '';
    
    if (!startYear || !startMonth || !endYear || !endMonth) {
        alert('조회기간을 모두 선택해주세요.');
        return;
    }
    
    // period 문자열 생성: "2025-01 ~ 2025-12"
    const period = `${startYear}-${startMonth.padStart(2, '0')} ~ ${endYear}-${endMonth.padStart(2, '0')}`;
    
    // periodInput에도 업데이트 (다운로드 등에서 사용)
    if (periodInput) {
        periodInput.value = period;
    }
    
    // 조회기간 형식 검증
    const periodRegex = /^\d{4}-\d{2}\s*~\s*\d{4}-\d{2}$/;
    if (!periodRegex.test(period)) {
        alert('조회기간 형식을 올바르게 입력해주세요. (예: 2024-01 ~ 2024-12)');
        return;
    }
    
    // 기간 파싱
    const parsedPeriod = parsePeriod(period);
    if (!parsedPeriod) {
        alert('조회기간을 올바르게 입력해주세요.');
        return;
    }
    
    // 로딩 표시
    queryBtn.textContent = '조회 중...';
    queryBtn.disabled = true;
    
    // 현재 사용자 정보 가져오기
    const user = getCurrentUser();
    if (!user) {
        alert('로그인이 필요합니다. Microsoft 365로 로그인해주세요.');
        queryBtn.textContent = '조회';
        queryBtn.disabled = false;
        // 로그인 오버레이 표시
        showAuthOverlay();
        hideApp();
        return;
    }
    
    // API_BASE_URL 확인
    if (!API_BASE_URL) {
        alert('❌ 잘못된 접속 방법입니다.\n\n이 애플리케이션은 서버를 통해 접속해야 합니다.\n\n✅ 올바른 접속 방법:\n1. "start-all.cmd" 파일을 실행하세요\n2. 자동으로 브라우저가 열립니다\n3. 또는 서버를 실행한 후 http://서버IP:3000 으로 접속하세요\n\n⚠️ 파일을 직접 열거나 file:// 프로토콜로는 작동하지 않습니다!');
        queryBtn.textContent = '조회';
        queryBtn.disabled = false;
        return;
    }
    
    // 서버 상태 확인 및 자동 실행
    let serverRunning = false;
    if (ENABLE_SERVER_API) {
        try {
            // 타임아웃을 위한 AbortController 사용
            const controller = new AbortController();
            const timeoutId = setTimeout(() => controller.abort(), 2000); // 2초 타임아웃
            
            console.log('🔍 서버 상태 확인:', `${API_BASE_URL}/api/health`);
            const healthCheck = await fetch(`${API_BASE_URL}/api/health`, { 
                method: 'GET',
                signal: controller.signal
            });
            
            clearTimeout(timeoutId);
            
            if (healthCheck.ok) {
                const healthResult = await healthCheck.json();
                if (healthResult.success) {
                    serverRunning = true;
                    console.log('✅ 서버가 실행 중입니다.');
                }
            } else {
                console.warn('⚠️ 서버 헬스체크 응답 오류:', healthCheck.status, healthCheck.statusText);
            }
        } catch (error) {
            // 헬스체크 실패 - 서버가 실행되지 않았을 수 있음
            console.log('⚠️ 서버 헬스체크 실패:', error.message);
            console.log('📍 현재 URL:', window.location.href);
            console.log('📍 API 베이스 URL:', API_BASE_URL);
            
            // 서버가 실행되지 않았으면 자동으로 시작 시도
            queryBtn.textContent = '서버 확인 중...';
            const serverStarted = await tryStartServer();
            
            if (serverStarted) {
                serverRunning = true;
                console.log('✅ 서버가 시작되었습니다. 조회를 계속합니다.');
            } else {
                // 서버 시작 실패 또는 사용자가 취소
                queryBtn.textContent = '조회';
                queryBtn.disabled = false;
                return;
            }
        }
    }

    // 서버에 요청 보내기
    try {
        let serverData = {
            settled: { monthly: [], detail: [] },
            unsettled: { amount: 0, detail: [] }
        };

        if (ENABLE_SERVER_API) {
            if (!API_BASE_URL) {
                throw new Error('API 베이스 URL이 없습니다. 서버를 통해 접속해야 합니다.');
            }
            
            try {
                const apiUrl = `${API_BASE_URL}/api/personal-settlement?period=${encodeURIComponent(period)}&userName=${encodeURIComponent(user.name)}`;
                console.log('📡 API 호출:', apiUrl);
                
                const response = await fetch(apiUrl);
                
                console.log('📡 API 응답 상태:', response.status, response.statusText);
                
                if (!response.ok) {
                    throw new Error(`서버 응답 오류: ${response.status} ${response.statusText}`);
                }
                
                const result = await response.json();
                
                if (!result.success) {
                    throw new Error(result.error || '데이터 조회에 실패했습니다.');
                }

                serverData = result.data || serverData;
                console.log('✅ 정산 데이터 로드 완료');
            } catch (apiError) {
                console.error('❌ 서버 API 호출 실패:', apiError);
                console.error('📍 현재 URL:', window.location.href);
                console.error('📍 API 베이스 URL:', API_BASE_URL);
                
                const currentUrl = window.location.href;
                const isFileProtocol = currentUrl.startsWith('file://');
                
                // file:// 프로토콜이 아니면 서버 자동 실행 시도
                if (!isFileProtocol && !serverRunning) {
                    queryBtn.textContent = '서버 확인 중...';
                    const serverStarted = await tryStartServer();
                    
                    if (serverStarted) {
                        // 서버가 시작되었으면 serverRunning 업데이트
                        serverRunning = true;
                        
                        // 서버가 시작되었으면 API 호출 재시도
                        console.log('✅ 서버가 시작되었습니다. API 호출을 재시도합니다.');
                        try {
                            const apiUrl = `${API_BASE_URL}/api/personal-settlement?period=${encodeURIComponent(period)}&userName=${encodeURIComponent(user.name)}`;
                            console.log('📡 API 재호출:', apiUrl);
                            
                            const response = await fetch(apiUrl);
                            
                            if (!response.ok) {
                                throw new Error(`서버 응답 오류: ${response.status} ${response.statusText}`);
                            }
                            
                            const result = await response.json();
                            
                            if (!result.success) {
                                throw new Error(result.error || '데이터 조회에 실패했습니다.');
                            }
                            
                            serverData = result.data || serverData;
                            console.log('✅ 정산 데이터 로드 완료 (재시도 성공)');
                        } catch (retryError) {
                            console.error('❌ API 재호출 실패:', retryError);
                            serverRunning = false; // 재시도 실패 시 false로 설정
                            // 재시도 실패 시 아래 에러 메시지 표시
                        }
                    } else {
                        // 서버 시작 실패 또는 사용자가 취소
                        queryBtn.textContent = '조회';
                        queryBtn.disabled = false;
                        return;
                    }
                }
                
                // 서버가 여전히 실행되지 않았거나 file:// 프로토콜인 경우
                if (!serverRunning) {
                    let errorMsg = '⚠️ 서버에 연결할 수 없습니다.\n\n';
                    
                    if (isFileProtocol) {
                        errorMsg += '❌ 파일을 직접 열었습니다.\n';
                        errorMsg += '이 애플리케이션은 서버를 통해 접속해야 합니다.\n\n';
                        errorMsg += '✅ 올바른 접속 방법:\n';
                        errorMsg += '1. 서버를 실행한 컴퓨터에서 "start-all.cmd" 파일 실행\n';
                        errorMsg += '2. 자동으로 브라우저가 열립니다\n';
                        errorMsg += '3. 다른 사람은 서버 IP 주소로 접속 (예: http://192.168.x.x:3000)\n\n';
                        errorMsg += '⚠️ 파일을 직접 열거나 file:// 프로토콜로는 작동하지 않습니다!';
                    } else {
                        errorMsg += '서버를 실행해야 데이터를 조회할 수 있습니다.\n\n';
                        errorMsg += '✅ 서버 실행 방법:\n';
                        errorMsg += '1. 서버를 실행한 컴퓨터에서 "start-all.cmd" 파일 실행\n';
                        errorMsg += '2. 서버 창에서 네트워크 IP 주소 확인 (예: http://192.168.x.x:3000)\n';
                        errorMsg += '3. 다른 사람은 그 주소로 접속\n\n';
                        errorMsg += '💡 현재 접속 주소: ' + currentUrl + '\n';
                        errorMsg += '💡 API 호출 주소: ' + (API_BASE_URL || '없음') + '\n\n';
                        errorMsg += '서버가 실행되면 이 페이지를 새로고침하고 다시 조회해주세요.';
                    }
                    
                    alert(errorMsg);
                    queryBtn.textContent = '조회';
                    queryBtn.disabled = false;
                    return;
                }
            }
        } else {
            throw new Error('서버 API가 비활성화되어 있습니다.');
        }
        
        // 기간에 맞는 데이터 필터링
        const parsedPeriod = parsePeriod(period);
        
        // 🔍 디버깅: 서버에서 받은 원본 데이터 확인
        const originalDetail = serverData.settled?.detail || [];
        console.log(`\n📊 [데이터 확인] 서버에서 받은 원본 데이터 개수: ${originalDetail.length}개`);
        
        // 🔍 디버깅: 조회기간에 2024가 포함된 경우 상세 확인
        if (parsedPeriod && (parsedPeriod.startYear === 2024 || parsedPeriod.endYear === 2024 || parsedPeriod.startYear <= 2024)) {
            // 정산월별 데이터 확인
            const byMonth = {};
            originalDetail.forEach(item => {
                const month = item.month || item.settlementMonth || '없음';
                if (!byMonth[month]) {
                    byMonth[month] = [];
                }
                byMonth[month].push(item);
            });
            console.log(`📊 [원본 데이터] 정산월별 개수:`, Object.keys(byMonth).sort().map(m => `${m}: ${byMonth[m].length}개`).join(', '));
            
            // 2024-12 정산월 데이터 확인
            const month2024_12 = originalDetail.filter(item => {
                const month = item.month || item.settlementMonth || '';
                return month === '2024-12';
            });
            console.log(`📊 [원본 데이터] 정산월 2024-12인 데이터: ${month2024_12.length}개`);
            if (month2024_12.length > 0) {
                console.log(`   📋 샘플 (처음 5개):`, month2024_12.slice(0, 5).map(item => ({
                    정산월: item.month || item.settlementMonth,
                    지급일: item.paymentDate,
                    사용처: item.merchant,
                    금액: item.amount,
                    출처: item.isFromSQL ? 'SQL' : '엑셀'
                })));
            }
            
            // 지급일별 데이터 확인
            const byPaymentDate = {};
            originalDetail.forEach(item => {
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
            console.log(`📊 [원본 데이터] 지급일(YYYY-MM)별 개수:`, Object.keys(byPaymentDate).sort().map(d => `${d}: ${byPaymentDate[d].length}개`).join(', '));
            
            // 2024-12 지급일 데이터 확인
            const payment2024_12 = originalDetail.filter(item => {
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
            console.log(`📊 [원본 데이터] 지급일 2024-12인 데이터: ${payment2024_12.length}개`);
            if (payment2024_12.length > 0) {
                console.log(`   📋 샘플 (처음 5개):`, payment2024_12.slice(0, 5).map(item => ({
                    정산월: item.month || item.settlementMonth,
                    지급일: item.paymentDate,
                    사용처: item.merchant,
                    금액: item.amount
                })));
            }
            
            // 2025-01 지급일 데이터 확인
            const payment2025_01 = originalDetail.filter(item => {
                if (!item.paymentDate) return false;
                const paymentDateStr = String(item.paymentDate).trim();
                let paymentYearMonth = '';
                if (/^\d{4}-\d{2}-\d{2}$/.test(paymentDateStr)) {
                    paymentYearMonth = paymentDateStr.substring(0, 7);
                } else if (/^\d{4}-\d{2}$/.test(paymentDateStr)) {
                    paymentYearMonth = paymentDateStr;
                }
                return paymentYearMonth === '2025-01';
            });
            console.log(`📊 [원본 데이터] 지급일 2025-01인 데이터: ${payment2025_01.length}개`);
            
            // 2025-02 지급일 데이터 확인
            const payment2025_02 = originalDetail.filter(item => {
                if (!item.paymentDate) return false;
                const paymentDateStr = String(item.paymentDate).trim();
                let paymentYearMonth = '';
                if (/^\d{4}-\d{2}-\d{2}$/.test(paymentDateStr)) {
                    paymentYearMonth = paymentDateStr.substring(0, 7);
                } else if (/^\d{4}-\d{2}$/.test(paymentDateStr)) {
                    paymentYearMonth = paymentDateStr;
                }
                return paymentYearMonth === '2025-02';
            });
            console.log(`📊 [원본 데이터] 지급일 2025-02인 데이터: ${payment2025_02.length}개`);
        }
        
        const filteredSettledDetail = parsedPeriod ? filterDataByPeriod(originalDetail, parsedPeriod) : originalDetail;
        
        // 🔍 디버깅: 필터링 후 데이터 확인
        console.log(`📊 [필터링 후] 데이터 개수: ${filteredSettledDetail.length}개`);
        
        // 🔍 디버깅: 조회기간에 2024가 포함된 경우 정산월별 개수 확인
        if (parsedPeriod && (parsedPeriod.startYear === 2024 || parsedPeriod.endYear === 2024 || parsedPeriod.startYear <= 2024)) {
            const filteredByMonth = {};
            filteredSettledDetail.forEach(item => {
                const month = item.month || item.settlementMonth || '없음';
                if (!filteredByMonth[month]) {
                    filteredByMonth[month] = [];
                }
                filteredByMonth[month].push(item);
            });
            console.log(`📊 [필터링 후] 정산월별 개수:`, Object.keys(filteredByMonth).sort().map(m => `${m}: ${filteredByMonth[m].length}개`).join(', '));
            
            // 2024-12 데이터 확인
            if (filteredByMonth['2024-12']) {
                console.log(`✅ [필터링 후] 2024-12 데이터: ${filteredByMonth['2024-12'].length}개`);
                if (filteredByMonth['2024-12'].length > 0) {
                    console.log(`   📋 첫 번째 2024-12 데이터:`, filteredByMonth['2024-12'][0]);
                }
            } else {
                console.warn(`⚠️ [필터링 후] 2024-12 데이터 없음`);
                
                // 원본 데이터에서 2024-12 확인
                const original2024_12 = originalDetail.filter(item => {
                    const month = item.month || item.settlementMonth || '';
                    return month === '2024-12';
                });
                console.log(`📊 [원본 데이터] 2024-12 데이터: ${original2024_12.length}개`);
                if (original2024_12.length > 0) {
                    console.log(`   📋 첫 번째 2024-12 원본 데이터:`, original2024_12[0]);
                    // 필터링 테스트
                    const testFiltered = filterDataByPeriod(original2024_12, parsedPeriod);
                    console.log(`   🔍 2024-12 데이터 필터링 테스트: ${testFiltered.length}개 (원본: ${original2024_12.length}개)`);
                    if (testFiltered.length === 0 && original2024_12.length > 0) {
                        console.error(`   ❌ [오류] 2024-12 데이터가 필터링에서 제외됨!`);
                        console.error(`   📅 조회기간: ${parsedPeriod.startYear}-${String(parsedPeriod.startMonth).padStart(2, '0')} ~ ${parsedPeriod.endYear}-${String(parsedPeriod.endMonth).padStart(2, '0')}`);
                        original2024_12.slice(0, 3).forEach((item, idx) => {
                            const month = item.month || item.settlementMonth || '';
                            const isInRange = isMonthInRange(month, parsedPeriod);
                            console.error(`   [${idx + 1}] 정산월="${month}", 포함=${isInRange}`);
                        });
                    }
                }
            }
        }
        
        // 🔥 월별 정산 요약: 상세 내역 데이터를 정산월 기준으로 집계
        const filteredMonthlyData = calculateMonthlySummaryFromDetail(filteredSettledDetail);
        
        // 미정산 데이터 별도 로드 (사용자 이름으로 필터링)
        let unsettledDetailData = [];
        if (API_BASE_URL) {
            try {
                const unsettledUrl = `${API_BASE_URL}/api/unsettled-data?userName=${encodeURIComponent(user.name)}`;
                console.log('📡 미정산 API 호출:', unsettledUrl);
                const unsettledResponse = await fetch(unsettledUrl);
            console.log('📡 미정산 API 응답 상태:', unsettledResponse.status, unsettledResponse.statusText);
            if (unsettledResponse.ok) {
                const unsettledResult = await unsettledResponse.json();
                console.log('📡 미정산 API 응답 데이터(raw):', unsettledResult);
                
                // 🔥 다양한 응답 구조 지원
                if (unsettledResult.success && unsettledResult.data && unsettledResult.data.unsettled) {
                    // 표준 응답 구조: { success: true, data: { unsettled: { detail: [...] } } }
                    unsettledDetailData = unsettledResult.data.unsettled.detail || [];
                    console.log(`✅ 미정산 데이터 로드 완료 (표준 구조) (${user.name}):`, unsettledDetailData.length, '개 항목');
                } else if (unsettledResult.data && unsettledResult.data.unsettled) {
                    // data.unsettled 구조
                    unsettledDetailData = unsettledResult.data.unsettled.detail || [];
                    console.log(`✅ 미정산 데이터 로드 완료 (data.unsettled) (${user.name}):`, unsettledDetailData.length, '개 항목');
                } else if (unsettledResult.unsettled) {
                    // unsettled 직접 구조
                    unsettledDetailData = unsettledResult.unsettled.detail || [];
                    console.log(`✅ 미정산 데이터 로드 완료 (unsettled 직접) (${user.name}):`, unsettledDetailData.length, '개 항목');
                } else if (Array.isArray(unsettledResult.data)) {
                    // 배열 직접 응답
                    unsettledDetailData = unsettledResult.data;
                    console.log(`✅ 미정산 데이터 로드 완료 (배열형 응답) (${user.name}):`, unsettledDetailData.length, '개 항목');
                } else {
                    console.warn('⚠️ 미정산 API 응답 구조를 인식할 수 없습니다:', unsettledResult);
                    console.warn('   응답 구조:', JSON.stringify(unsettledResult, null, 2));
                }
                
                if (unsettledDetailData.length > 0) {
                    console.log('📋 첫 번째 미정산 데이터:', unsettledDetailData[0]);
                } else {
                    console.warn('⚠️ 미정산 데이터가 비어있습니다.');
                }
            } else {
                const errorText = await unsettledResponse.text();
                throw new Error(`미정산 API 응답 오류: ${unsettledResponse.status} ${unsettledResponse.statusText} / ${errorText}`);
            }
            } catch (unsettledError) {
                console.error('❌ 미정산 데이터 로드 오류:', unsettledError);
                // 미정산 데이터 로드 실패해도 계속 진행 (정산 데이터는 표시)
                console.warn('⚠️ 미정산 데이터를 로드할 수 없지만 계속 진행합니다.');
            }
        } else {
            console.warn('⚠️ API_BASE_URL이 없어 미정산 데이터를 로드할 수 없습니다.');
        }
        
        // 🔥 미정산 상세내역은 조회 기간에 상관없이 SQL에서 가져온 데이터를 그대로 사용
        const filteredUnsettledDetail = unsettledDetailData;

        // 🔥 월별 정산 요약은 항상 상세 내역에서 계산하므로 서버의 monthly 데이터는 사용하지 않음
        latestServerData = {
            settled: {
                monthly: [], // 서버의 monthly 데이터는 사용하지 않음 (상세 내역에서 계산)
                detail: serverData.settled?.detail || []
            },
            unsettled: {
                amount: serverData.unsettled?.amount || 0,
                detail: unsettledDetailData || []
            }
        };

        selectedMonth = null;
        currentFilteredMonthlyData = filteredMonthlyData;
        currentFilteredSettledDetail = filteredSettledDetail;
        currentFilteredUnsettledDetail = filteredUnsettledDetail;
        console.log('📋 필터링된 미정산 데이터:', filteredUnsettledDetail.length, '개 항목');
        
        if (filteredUnsettledDetail.length > 0) {
          console.log('📋 미정산 데이터 샘플 (처음 3개):', filteredUnsettledDetail.slice(0, 3));
        } else if (unsettledDetailData.length > 0) {
          console.warn('⚠️ 미정산 데이터가 있지만 표시되지 않습니다.');
          console.log('   - 원본 데이터:', unsettledDetailData.length, '개');
          console.log('   - 첫 번째 원본 데이터:', unsettledDetailData[0]);
        } else {
          console.warn('⚠️ 미정산 데이터가 없습니다.');
        }
        
        console.log('조회 기간:', parsedPeriod);
        console.log('사용자:', user.name);
        console.log('필터링된 월별 데이터:', filteredMonthlyData);
        console.log('필터링된 정산 상세:', filteredSettledDetail);
        console.log('필터링된 미정산 상세:', filteredUnsettledDetail);
        
        // 정산월 데이터 확인 (첫 번째 항목)
        if (filteredSettledDetail && filteredSettledDetail.length > 0) {
            console.log('🔍 첫 번째 정산 상세 항목:', filteredSettledDetail[0]);
            console.log('🔍 정산월 값:', filteredSettledDetail[0].month);
        }
        
        // 선택된 월이 있으면 해당 월로 추가 필터링
        const activeTabInput = document.querySelector('.tab-item.active input[type="radio"]');
        const activeTab = activeTabInput ? activeTabInput.value : 'settled';
        if (activeTab === 'settled') {
            // 🔥 월별 정산 요약: 항상 현재 상세 내역에서 계산 (함수 내부에서 처리)
            // filteredMonthlyData는 참고용으로만 저장하고, 실제 표시는 함수 내부에서 계산
            currentFilteredMonthlyData = filteredMonthlyData;
            updateMonthlySummary();
            const detailData = selectedMonth ? 
                filteredSettledDetail.filter(item => {
                    // 정산월 컬럼에서 가져온 month 값 사용
                    return item.month === selectedMonth;
                }) : 
                filteredSettledDetail;
            updateSettledDetail(detailData);
        } else {
            updateMonthlySummary([]);
            updateSettledDetail([]);
        }
        
        // 미정산 데이터 업데이트
        console.log('🔄 updateUnsettledDetail 호출 전:', filteredUnsettledDetail.length, '개 항목');
        updateUnsettledDetail(filteredUnsettledDetail);
        console.log('✅ updateUnsettledDetail 호출 완료');
        
        // 미정산 금액 재계산
        const unsettledAmount = calculateUnsettledAmount(filteredUnsettledDetail);
        const unsettledTotalValue = document.getElementById('unsettled-total-value');
        if (unsettledTotalValue) {
            unsettledTotalValue.textContent = formatNumber(unsettledAmount);
        }
        console.log('💰 미정산 금액:', unsettledAmount);

        serverData.settled = serverData.settled || {};
        serverData.unsettled = serverData.unsettled || {};
        serverData.settled.detail = filteredSettledDetail;
        serverData.unsettled.detail = filteredUnsettledDetail;
        

        
        
        
        
        
        console.log("📌 AI로 전달되는 serverData:", JSON.parse(JSON.stringify(serverData)));
        console.log("📌 settled.detail:", serverData?.settled?.detail);
        console.log("📌 unsettled.detail:", serverData?.unsettled?.detail);

        // 상단 미정산 금액 테이블 업데이트
        updateUnsettledSummaryTable(filteredUnsettledDetail, unsettledAmount);
        
        alert(`조회기간: ${period}에 대한 데이터를 조회했습니다. (사용자: ${user.name})`);
    } catch (error) {
        console.error('데이터 조회 중 오류:', error);
        alert('데이터 조회 중 오류가 발생했습니다: ' + error.message);
    } finally {
        // 버튼 상태 복원
        queryBtn.textContent = '조회';
        queryBtn.disabled = false;
    }
});


// 월별 정산 요약 다운로드 버튼 클릭 이벤트
monthlySummaryDownloadBtn.addEventListener('click', () => {
    // 다운로드 로딩 표시
    monthlySummaryDownloadBtn.textContent = '다운로드 중...';
    monthlySummaryDownloadBtn.disabled = true;
    
    try {
        const period = periodInput.value.trim();
        const dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const filename = period 
            ? `월별정산요약_${period.replace(/\s/g, '')}.xlsx`
            : `월별정산요약_${dateStr}.xlsx`;
        downloadTableAsExcel('monthly-summary-table', filename);
    } catch (error) {
        alert('다운로드 중 오류가 발생했습니다: ' + error.message);
    } finally {
        setTimeout(() => {
            monthlySummaryDownloadBtn.textContent = '다운로드';
            monthlySummaryDownloadBtn.disabled = false;
        }, 1000);
    }
});

// 월정산 상세내역 다운로드 버튼 클릭 이벤트
settledDownloadBtn.addEventListener('click', () => {
    // 다운로드 로딩 표시
    settledDownloadBtn.textContent = '다운로드 중...';
    settledDownloadBtn.disabled = true;
    
    try {
        const period = periodInput.value.trim();
        const dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const filename = period 
            ? `월정산상세내역_${period.replace(/\s/g, '')}.xlsx`
            : `월정산상세내역_${dateStr}.xlsx`;
        downloadTableAsExcel('settled-detail-table', filename);
    } catch (error) {
        alert('다운로드 중 오류가 발생했습니다: ' + error.message);
    } finally {
        // 버튼 상태 복원
        setTimeout(() => {
            settledDownloadBtn.textContent = '다운로드';
            settledDownloadBtn.disabled = false;
        }, 1000);
    }
});

// 미정산 상세내역 다운로드 버튼 클릭 이벤트
unsettledDownloadBtn.addEventListener('click', () => {
    // 다운로드 로딩 표시
    unsettledDownloadBtn.textContent = '다운로드 중...';
    unsettledDownloadBtn.disabled = true;
    
    try {
        const period = periodInput.value.trim();
        const dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '');
        const filename = period 
            ? `미정산상세내역_${period.replace(/\s/g, '')}.xlsx`
            : `미정산상세내역_${dateStr}.xlsx`;
        downloadTableAsExcel('unsettled-detail-table', filename);
    } catch (error) {
        alert('다운로드 중 오류가 발생했습니다: ' + error.message);
    } finally {
        // 버튼 상태 복원
        setTimeout(() => {
            unsettledDownloadBtn.textContent = '다운로드';
            unsettledDownloadBtn.disabled = false;
        }, 1000);
    }
});

// 페이지 로드 시 초기 데이터 설정
document.addEventListener('DOMContentLoaded', async () => {
    // 서버 상태 확인 (자동 시작 기능 비활성화 - 수동 시작만 허용)
    if (ENABLE_SERVER_API && API_BASE_URL) {
        console.log('🔍 페이지 로드 시 서버 상태 확인 중...');
        const serverRunning = await checkServerStatus();
        
        if (!serverRunning) {
            console.log('⚠️ 서버가 실행 중이지 않습니다.');
            console.log('💡 서버를 시작하려면:');
            console.log('   1. 프로젝트 폴더에서 "start-all.cmd" 파일을 더블클릭하세요');
            console.log('   2. 또는 터미널에서 "node server.js" 명령어를 실행하세요');
            console.log('   3. 서버가 시작되면 이 페이지를 새로고침하세요');
        } else {
            console.log('✅ 서버가 이미 실행 중입니다.');
            // 서버가 이미 실행 중이면 시트 목록 로드
        }
    } else {
        // API_BASE_URL이 없는 경우 (file:// 프로토콜 등)
        console.warn('⚠️ API_BASE_URL이 없습니다. 서버를 통해 접속해야 합니다.');
    }
    
    // M365 인증 초기화 및 가드 (비활성화 시 자동 통과)
    await initializeMsalAndGuard();

    if (ENABLE_AUTH_GUARD && m365LoginBtn) {
        m365LoginBtn.addEventListener('click', () => loginWithM365());
    }

    // 사용자 정보 표시 초기화
    updateUserDisplay();

    // 현재 날짜를 기본값으로 설정
    const now = new Date();
    const currentYear = now.getFullYear();
    const currentMonth = String(now.getMonth() + 1).padStart(2, '0');
    const defaultPeriod = `${currentYear}-${currentMonth} ~ ${currentYear}-${currentMonth}`;
    periodInput.value = defaultPeriod;
    
});

// 키보드 이벤트 처리
document.addEventListener('keydown', (e) => {
    // Enter 키로 조회 실행
    if (e.key === 'Enter' && e.target === periodInput) {
        queryBtn.click();
    }
    
    // ESC 키로 입력 필드 초기화
    if (e.key === 'Escape' && e.target === periodInput) {
        periodInput.value = '';
        periodInput.focus();
    }
});

// 테이블 행 호버 효과
function addTableHoverEffects() {
    const tables = document.querySelectorAll('table tbody');
    tables.forEach(tbody => {
        const rows = tbody.querySelectorAll('tr');
        rows.forEach(row => {
            row.addEventListener('mouseenter', () => {
                row.style.backgroundColor = '#e8f4fd';
            });
            row.addEventListener('mouseleave', () => {
                row.style.backgroundColor = '';
            });
        });
    });
}

// 윈도우 리사이즈 이벤트
window.addEventListener('resize', () => {
    // 모바일에서 테이블 스크롤 최적화
    const summaryTbody = document.querySelector('.summary-table tbody');
    const detailTbodies = document.querySelectorAll('.detail-table tbody');
    
    if (window.innerWidth <= 768) {
        if (summaryTbody) summaryTbody.style.maxHeight = '200px';
        detailTbodies.forEach(tbody => {
            tbody.style.maxHeight = '200px';
        });
    } else {
        if (summaryTbody) summaryTbody.style.maxHeight = '300px';
        detailTbodies.forEach(tbody => {
            tbody.style.maxHeight = '300px';
        });
    }
});

// 초기화 함수
function initializeApp() {
    addTableHoverEffects();
    
    // 테이블 스크롤 최적화
    const summaryTbody = document.querySelector('.summary-table tbody');
    const detailTbodies = document.querySelectorAll('.detail-table tbody');
    
    if (summaryTbody) {
        summaryTbody.style.maxHeight = window.innerWidth <= 768 ? '200px' : '300px';
    }
    
    detailTbodies.forEach(tbody => {
        tbody.style.maxHeight = window.innerWidth <= 768 ? '200px' : '300px';
    });
}


// 관리자 테이블 데이터를 엑셀로 다운로드
function downloadAdminTableAsExcel() {
    const thead = document.getElementById('excel-data-thead');
    const tbody = document.getElementById('excel-data-tbody');
    
    if (!thead || !tbody) {
        alert('테이블을 찾을 수 없습니다.');
        return;
    }
    
    // 헤더 데이터 추출
    const headerRows = thead.querySelectorAll('tr');
    const data = [];
    
    headerRows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('th');
        cells.forEach(cell => {
            rowData.push(cell.textContent.trim());
        });
        if (rowData.length > 0) {
            data.push(rowData);
        }
    });
    
    // 본문 데이터 추출
    const bodyRows = tbody.querySelectorAll('tr');
    bodyRows.forEach(row => {
        const rowData = [];
        const cells = row.querySelectorAll('td');
        cells.forEach(cell => {
            const text = cell.textContent.trim();
            // 금액 형식이면 숫자로 변환하여 삽입
            const num = parseCurrencyToNumber(text);
            rowData.push(num !== null ? num : text);
        });
        if (rowData.length > 0) {
            data.push(rowData);
        }
    });
    
    if (data.length === 0) {
        alert('다운로드할 데이터가 없습니다.');
        return;
    }
    
    // 워크북 생성
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    
    // 컬럼 너비 설정
    const colWidths = [];
    if (data[0]) {
        data[0].forEach((_, index) => {
            colWidths.push({ wch: 15 });
        });
    }
    ws['!cols'] = colWidths;
    
    // 시트명을 현재 선택된 시트명으로 설정
    const sheetName = currentSheetName || 'Sheet1';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    // 파일명 생성 (시트명 포함)
    const filename = `관리자_데이터_${sheetName}_${new Date().toISOString().split('T')[0]}.xlsx`;
    
    // 파일 다운로드
    XLSX.writeFile(wb, filename);
    console.log(`✅ 엑셀 다운로드 완료: ${filename}`);
}

// 관리자 화면 - 엑셀 데이터 로드 및 표시
let excelDataCache = null; // 전체 데이터 캐시
let filteredDataCache = null; // 필터링된 데이터
let currentSheetName = null;
let currentPage = 1;
let currentPageSize = 50; // 관리자 화면은 항상 50개씩 표시
let totalRows = 0;
let totalPages = 1;
let currentHeaders = [];
let currentFilterConditions = {
    searchTerm: ''
};
let isAdminMode = true; // 관리자 화면에서는 항상 true (모든 데이터 표시)

// 엑셀 시트 목록 로드
async function loadExcelSheets() {
    if (!API_BASE_URL) {
        console.error('❌ API_BASE_URL이 없습니다. 서버를 통해 접속해야 합니다.');
        return;
    }
    
    // 서버 상태 먼저 확인
    console.log('🔍 시트 목록 로드 전 서버 상태 확인 중...');
    let serverRunning = await checkServerStatus();
    
    // 서버가 실행되지 않은 경우 안내만 하고 중단
    if (!serverRunning) {
        console.log('⚠️ 서버가 실행 중이지 않습니다. 서버를 먼저 실행해주세요.');
        console.log('💡 방법: start-all.cmd 실행 또는 터미널에서 node server.js');
        return;
    }
    
    try {
        console.log('📡 시트 목록 API 호출:', `${API_BASE_URL}/api/sheets`);
        const response = await fetch(`${API_BASE_URL}/api/sheets`);
        
        // 응답 상태 확인
        if (!response.ok) {
            const contentType = response.headers.get('content-type');
            let errorMessage = `서버 오류 (${response.status} ${response.statusText})`;
            
            // HTML이 반환된 경우 (서버가 실행되지 않았거나 잘못된 경로)
            if (contentType && contentType.includes('text/html')) {
                errorMessage = '서버가 실행되지 않았거나 API 엔드포인트를 찾을 수 없습니다. 서버가 실행 중인지 확인해주세요.';
            } else {
                // JSON 에러 메시지 시도
                try {
                    const errorData = await response.json();
                    errorMessage = errorData.error || errorMessage;
                } catch (e) {
                    // JSON 파싱 실패 시 기본 메시지 사용
                }
            }
            
            throw new Error(errorMessage);
        }
        
        // Content-Type 확인
        const contentType = response.headers.get('content-type');
        if (!contentType || !contentType.includes('application/json')) {
            throw new Error('서버가 JSON 형식이 아닌 응답을 반환했습니다. 서버가 정상적으로 실행 중인지 확인해주세요.');
        }
        
        const result = await response.json();
        
        if (result.success && result.sheets) {
            const sheetSelect = document.getElementById('excel-sheet-select');
            if (sheetSelect) {
                // 기존 옵션 제거 (첫 번째 옵션 제외)
                while (sheetSelect.children.length > 1) {
                    sheetSelect.removeChild(sheetSelect.lastChild);
                }
                
                // 시트 목록 추가 (2024, 2025 등)
                result.sheets.forEach(sheetName => {
                    const option = document.createElement('option');
                    option.value = sheetName;
                    option.textContent = sheetName;
                    sheetSelect.appendChild(option);
                });
                
                console.log(`✅ 시트 목록 로드 완료: ${result.sheets.length}개 시트 (${result.sheets.join(', ')})`);
                
                // 시트가 있으면 첫 번째 시트 자동 선택 (2025 우선, 없으면 첫 번째)
                if (result.sheets.length > 0) {
                    // 2025가 있으면 2025를, 없으면 첫 번째 시트를 선택
                    const sheet2025 = result.sheets.find(s => s === '2025');
                    const defaultSheet = sheet2025 || result.sheets[0];
                    sheetSelect.value = defaultSheet;
                    currentPage = 1; // 첫 페이지로 리셋
                    currentFilterConditions.searchTerm = ''; // 필터 초기화
                    // 검색 입력 필드 초기화
                    const searchInput = document.getElementById('excel-search-input');
                    if (searchInput) {
                        searchInput.value = '';
                    }
                    loadExcelData(defaultSheet, false);
                }
            } else {
                console.warn('⚠️ 시트 선택 드롭다운을 찾을 수 없습니다.');
            }
        } else {
            console.error('❌ 시트 목록 로드 실패:', result.error || '알 수 없는 오류');
            alert('시트 목록을 불러올 수 없습니다: ' + (result.error || '알 수 없는 오류'));
        }
    } catch (error) {
        console.error('❌ 시트 목록 로드 오류:', error);
        
        // 네트워크 오류인 경우 서버 상태 재확인
        if (error.message.includes('Failed to fetch') || error.message.includes('NetworkError')) {
            console.log('🔍 네트워크 오류 발생. 서버 상태 재확인 중...');
            const serverRunning = await checkServerStatus();
            
            if (!serverRunning) {
                alert('서버가 실행되지 않았습니다.\n\n프로젝트 폴더에서 "start-all.cmd" 파일을 실행해주세요.');
                return;
            }
        }
        
        alert('시트 목록을 불러오는 중 오류가 발생했습니다: ' + error.message);
    }
}

// 엑셀 데이터 로드 (전체 데이터 가져오기)
async function loadExcelData(sheetName, forceReload = false) {
    if (!sheetName) {
        return;
    }
    
    // 페이지 크기를 항상 50으로 고정
    currentPageSize = 50;
    
    // 같은 시트이고 이미 캐시가 있으면 재로드하지 않음
    if (!forceReload && currentSheetName === sheetName && excelDataCache && excelDataCache.length > 0) {
        console.log('✅ 캐시된 데이터 사용');
        // 캐시된 데이터를 사용하되, 반드시 필터링 적용
        // 검색어가 없으면 빈 테이블 표시
        applyFiltersAndRender(1);
        return;
    }
    
    if (!API_BASE_URL) {
        console.error('❌ API_BASE_URL이 없습니다. 서버를 통해 접속해야 합니다.');
        return;
    }
    
    try {
        // 전체 데이터 가져오기 (페이지네이션 없이)
        const apiUrl = `${API_BASE_URL}/api/data/${encodeURIComponent(sheetName)}?page=1&limit=999999`;
        console.log('📡 엑셀 데이터 API 호출:', apiUrl);
        const response = await fetch(apiUrl);
        const result = await response.json();
        
        if (result.success) {
            excelDataCache = result.data; // 전체 데이터 캐시
            currentSheetName = sheetName;
            currentHeaders = result.headers;
            totalRows = result.totalRows; // 전체 행 수
            
            console.log(`✅ 전체 데이터 로드 완료: ${totalRows}개 행 (캐시에 저장됨)`);
            if (isAdminMode) {
                console.log(`👑 관리자 모드: 모든 데이터가 표시됩니다.`);
            } else {
                console.log(`⚠️ 검색어를 입력해야 데이터가 표시됩니다.`);
            }
            
            // 필터링 조건 강제 초기화 (검색어 없음)
            currentFilterConditions.searchTerm = '';
            
            // 검색 입력 필드도 강제 초기화
            const searchInput = document.getElementById('excel-search-input');
            if (searchInput) {
                searchInput.value = '';
            }
            
            // 관리자 모드: 초기 로드 시 모든 데이터 표시
            // 일반 사용자 모드: 검색어가 없으면 안내 메시지 표시
            if (!isAdminMode) {
                // 일반 사용자 모드: 테이블을 먼저 비우고 안내 메시지 표시
                const tbody = document.getElementById('excel-data-tbody');
                const thead = document.getElementById('excel-data-thead');
                if (tbody) tbody.innerHTML = '';
                if (thead) {
                    thead.innerHTML = '<tr><th style="text-align: center; padding: 20px;">검색어를 입력하여 데이터를 조회하세요.</th></tr>';
                }
            }
            
            // 필터링 적용 및 렌더링
            // 관리자 모드: 검색어 없이도 모든 데이터 표시
            // 일반 사용자 모드: 검색어가 없으면 안내 메시지만 표시
            applyFiltersAndRender(1);
        } else {
            alert('데이터 로드 실패: ' + (result.error || '알 수 없는 오류'));
        }
    } catch (error) {
        console.error('엑셀 데이터 로드 오류:', error);
        alert('데이터 로드 중 오류가 발생했습니다: ' + error.message);
    }
}

// 필터링 조건 적용 및 렌더링
function applyFiltersAndRender(page = 1) {
    const tbody = document.getElementById('excel-data-tbody');
    const thead = document.getElementById('excel-data-thead');
    
    // 데이터가 없으면 빈 테이블 표시
    if (!excelDataCache || excelDataCache.length === 0) {
        if (thead) {
            thead.innerHTML = '<tr><th style="text-align: center; padding: 20px;">데이터를 로드 중입니다...</th></tr>';
        }
        if (tbody) {
            tbody.innerHTML = '';
        }
        totalRows = 0;
        totalPages = 0;
        renderPagination();
        updatePageInfo();
        return;
    }
    
    // 관리자 모드: 검색어가 없어도 모든 데이터 표시
    // 일반 사용자 모드: 검색어가 없으면 데이터를 표시하지 않음
    const hasSearchTerm = currentFilterConditions.searchTerm && currentFilterConditions.searchTerm.trim() !== '';
    
    if (!isAdminMode && !hasSearchTerm) {
        // 일반 사용자 모드이고 검색어가 없으면 안내 메시지만 표시
        if (thead) {
            thead.innerHTML = '<tr><th style="text-align: center; padding: 20px;">검색어를 입력하여 데이터를 조회하세요.</th></tr>';
        }
        if (tbody) {
            tbody.innerHTML = '';
        }
        totalRows = 0;
        totalPages = 0;
        renderPagination();
        updatePageInfo();
        return;
    }
    
    // 필터링 적용
    // 관리자 모드: 검색어가 있으면 필터링, 없으면 전체 데이터
    // 일반 사용자 모드: 검색어가 있을 때만 필터링
    if (hasSearchTerm) {
        filteredDataCache = filterData(excelDataCache, currentFilterConditions);
    } else {
        // 관리자 모드이고 검색어가 없으면 전체 데이터 표시
        filteredDataCache = isAdminMode ? excelDataCache : [];
    }
    
    // 필터링된 데이터가 없으면 메시지 표시
    if (!filteredDataCache || filteredDataCache.length === 0) {
        // 검색어가 있지만 결과가 없으면 헤더는 표시하고 메시지 표시
        if (thead && currentHeaders && currentHeaders.length > 0) {
            thead.innerHTML = '';
            const headerRow = document.createElement('tr');
            currentHeaders.forEach(header => {
                const th = document.createElement('th');
                th.textContent = header || '';
                headerRow.appendChild(th);
            });
            thead.appendChild(headerRow);
        }
        if (tbody) {
            tbody.innerHTML = '';
            const row = document.createElement('tr');
            const cell = document.createElement('td');
            cell.colSpan = (currentHeaders && currentHeaders.length > 0) ? currentHeaders.length : 1;
            cell.style.textAlign = 'center';
            cell.style.padding = '20px';
            cell.textContent = '검색 결과가 없습니다. 다른 검색어를 시도해주세요.';
            row.appendChild(cell);
            tbody.appendChild(row);
        }
        totalRows = 0;
        totalPages = 0;
        renderPagination();
        updatePageInfo();
        return;
    }
    
    // 페이지네이션 적용
    const startIndex = (page - 1) * currentPageSize;
    const endIndex = startIndex + currentPageSize;
    const paginatedData = filteredDataCache.slice(startIndex, endIndex);
    
    // 페이지 정보 업데이트
    currentPage = page;
    totalRows = filteredDataCache.length;
    totalPages = Math.ceil(totalRows / currentPageSize);
    
    // 테이블 렌더링 (필터링된 데이터만)
    renderExcelDataTable(currentHeaders, paginatedData);
    renderPagination();
    updatePageInfo();
}

// 데이터 필터링 함수
function filterData(data, conditions) {
    if (!data || data.length === 0) {
        return [];
    }
    
    // 필터링 조건이 없으면 빈 배열 반환 (데이터를 표시하지 않음)
    // 사용자가 명시적으로 필터 조건을 설정해야 데이터가 표시됨
    if (!conditions.searchTerm || conditions.searchTerm.trim() === '') {
        // 검색어가 없으면 빈 배열 반환하여 데이터를 표시하지 않음
        return [];
    }
    
    // 검색어가 있으면 필터링 적용
    const searchLower = conditions.searchTerm.toLowerCase().trim();
    let filtered = data.filter(row => {
        return Object.values(row).some(value => 
            String(value || '').toLowerCase().includes(searchLower)
        );
    });
    
    // 추가 필터링 조건을 여기에 추가할 수 있습니다
    // 예: 특정 컬럼 값 필터링, 날짜 범위 필터링 등
    
    return filtered;
}

// 엑셀 데이터 테이블 렌더링 (필터링된 데이터만 표시)
function renderExcelDataTable(headers, data) {
    const thead = document.getElementById('excel-data-thead');
    const tbody = document.getElementById('excel-data-tbody');
    
    if (!thead || !tbody) {
        return;
    }
    
    // 데이터가 없으면 렌더링하지 않음 (applyFiltersAndRender에서 처리)
    if (!data || data.length === 0) {
        return;
    }
    
    // 헤더 렌더링
    thead.innerHTML = '';
    const headerRow = document.createElement('tr');
    headers.forEach(header => {
        const th = document.createElement('th');
        th.textContent = header || '';
        th.style.position = 'relative'; // 리사이즈 핸들을 위한 위치 설정
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    
    // 데이터 렌더링 (필터링된 데이터만)
    tbody.innerHTML = '';
    data.forEach(rowData => {
        const row = document.createElement('tr');
        headers.forEach(header => {
            const cell = document.createElement('td');
            const value = rowData[header] || '';
            // 숫자 형식인 경우 포맷팅
            if (typeof value === 'number') {
                cell.textContent = formatNumber(value);
            } else {
                cell.textContent = String(value);
            }
            row.appendChild(cell);
        });
        tbody.appendChild(row);
    });
    
    // 관리자 테이블의 모든 열에 리사이즈 핸들 추가
    setTimeout(() => {
        initializeAdminTableResize(headers.length);
    }, 100);
}

// 페이지네이션 UI 렌더링
function renderPagination() {
    const paginationDiv = document.getElementById('excel-pagination');
    if (!paginationDiv) {
        return;
    }
    
    paginationDiv.innerHTML = '';
    
    if (totalPages <= 1) {
        return; // 페이지가 1개 이하면 페이지네이션 표시 안 함
    }
    
    // 이전 버튼
    const prevBtn = document.createElement('button');
    prevBtn.className = 'btn btn-secondary btn-small';
    prevBtn.textContent = '이전';
    prevBtn.disabled = currentPage === 1;
    prevBtn.addEventListener('click', () => {
        if (currentPage > 1) {
            applyFiltersAndRender(currentPage - 1);
        }
    });
    paginationDiv.appendChild(prevBtn);
    
    // 페이지 번호 버튼들
    const maxButtons = 10; // 최대 표시할 페이지 버튼 수
    let startPage = Math.max(1, currentPage - Math.floor(maxButtons / 2));
    let endPage = Math.min(totalPages, startPage + maxButtons - 1);
    
    if (endPage - startPage < maxButtons - 1) {
        startPage = Math.max(1, endPage - maxButtons + 1);
    }
    
    // 첫 페이지 버튼
    if (startPage > 1) {
        const firstBtn = document.createElement('button');
        firstBtn.className = 'btn btn-secondary btn-small';
        firstBtn.textContent = '1';
        firstBtn.addEventListener('click', () => {
            applyFiltersAndRender(1);
        });
        paginationDiv.appendChild(firstBtn);
        
        if (startPage > 2) {
            const ellipsis = document.createElement('span');
            ellipsis.textContent = '...';
            ellipsis.style.padding = '0 5px';
            paginationDiv.appendChild(ellipsis);
        }
    }
    
    // 페이지 번호 버튼들
    for (let i = startPage; i <= endPage; i++) {
        const pageBtn = document.createElement('button');
        pageBtn.className = i === currentPage ? 'btn btn-primary btn-small' : 'btn btn-secondary btn-small';
        pageBtn.textContent = i;
        pageBtn.addEventListener('click', () => {
            applyFiltersAndRender(i);
        });
        paginationDiv.appendChild(pageBtn);
    }
    
    // 마지막 페이지 버튼
    if (endPage < totalPages) {
        if (endPage < totalPages - 1) {
            const ellipsis = document.createElement('span');
            ellipsis.textContent = '...';
            ellipsis.style.padding = '0 5px';
            paginationDiv.appendChild(ellipsis);
        }
        
        const lastBtn = document.createElement('button');
        lastBtn.className = 'btn btn-secondary btn-small';
        lastBtn.textContent = totalPages;
        lastBtn.addEventListener('click', () => {
            applyFiltersAndRender(totalPages);
        });
        paginationDiv.appendChild(lastBtn);
    }
    
    // 다음 버튼
    const nextBtn = document.createElement('button');
    nextBtn.className = 'btn btn-secondary btn-small';
    nextBtn.textContent = '다음';
    nextBtn.disabled = currentPage === totalPages;
    nextBtn.addEventListener('click', () => {
        if (currentPage < totalPages) {
            applyFiltersAndRender(currentPage + 1);
        }
    });
    paginationDiv.appendChild(nextBtn);
}

// 페이지 정보 업데이트
function updatePageInfo() {
    const pageInfo = document.getElementById('excel-page-info');
    if (pageInfo) {
        const startRow = totalRows === 0 ? 0 : (currentPage - 1) * currentPageSize + 1;
        const endRow = Math.min(currentPage * currentPageSize, totalRows);
        pageInfo.textContent = `전체 ${formatNumber(totalRows)}개 중 ${formatNumber(startRow)}-${formatNumber(endRow)}개 표시`;
    }
}

// 관리자 화면 이벤트 리스너 설정
function setupAdminExcelDataHandlers() {
    // 시트 선택 변경
    const sheetSelect = document.getElementById('excel-sheet-select');
    if (sheetSelect) {
        sheetSelect.addEventListener('change', (e) => {
            const selectedSheet = e.target.value;
            if (selectedSheet) {
                currentPage = 1; // 시트 변경 시 첫 페이지로
                currentFilterConditions.searchTerm = ''; // 필터 초기화
                const searchInput = document.getElementById('excel-search-input');
                if (searchInput) {
                    searchInput.value = '';
                }
                loadExcelData(selectedSheet, true); // 강제 재로드
            } else {
                const tbody = document.getElementById('excel-data-tbody');
                const thead = document.getElementById('excel-data-thead');
                const paginationDiv = document.getElementById('excel-pagination');
                if (tbody) {
                    tbody.innerHTML = '<tr><td colspan="100%" style="text-align: center; padding: 20px;">시트를 선택하세요.</td></tr>';
                }
                if (thead) {
                    thead.innerHTML = '<tr><th>시트를 선택하세요.</th></tr>';
                }
                if (paginationDiv) {
                    paginationDiv.innerHTML = '';
                }
                excelDataCache = null;
                filteredDataCache = null;
                updatePageInfo();
            }
        });
    }
    
    // 페이지 크기 변경
    const pageSizeSelect = document.getElementById('excel-page-size');
    if (pageSizeSelect) {
        pageSizeSelect.addEventListener('change', (e) => {
            const newPageSize = parseInt(e.target.value);
            currentPageSize = newPageSize;
            currentPage = 1; // 페이지 크기 변경 시 첫 페이지로
            if (currentSheetName && excelDataCache) {
                applyFiltersAndRender(1);
            }
        });
    }
    
    // 새로고침 버튼
    const refreshBtn = document.getElementById('refresh-excel-data-btn');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => {
            loadExcelSheets();
            if (currentSheetName) {
                currentPage = 1;
                currentFilterConditions.searchTerm = '';
                const searchInput = document.getElementById('excel-search-input');
                if (searchInput) {
                    searchInput.value = '';
                }
                loadExcelData(currentSheetName, true); // 강제 재로드
            }
        });
    }
    
    // 다운로드 버튼
    const downloadBtn = document.getElementById('excel-download-btn');
    if (downloadBtn) {
        downloadBtn.addEventListener('click', () => {
            if (!currentSheetName) {
                alert('시트를 먼저 선택해주세요.');
                return;
            }
            
            // 관리자 테이블 데이터 다운로드
            downloadAdminTableAsExcel();
        });
    }
    
    // 검색 기능 (전체 데이터에서 필터링)
    const searchBtn = document.getElementById('excel-search-btn');
    const searchInput = document.getElementById('excel-search-input');
    if (searchBtn && searchInput) {
        searchBtn.addEventListener('click', () => {
            if (!excelDataCache || excelDataCache.length === 0) {
                alert('먼저 시트를 선택하고 데이터를 로드해주세요.');
                return;
            }
            
            const searchTerm = searchInput.value.trim();
            currentFilterConditions.searchTerm = searchTerm;
            currentPage = 1; // 검색 시 첫 페이지로
            applyFiltersAndRender(1);
        });
        
        // Enter 키로 검색
        searchInput.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                searchBtn.click();
            }
        });
        
        // 실시간 검색 (입력 시마다 필터링)
        let searchTimeout = null;
        searchInput.addEventListener('input', (e) => {
            const searchTerm = e.target.value.trim();
            
            // 디바운싱: 300ms 후에 검색 실행
            clearTimeout(searchTimeout);
            searchTimeout = setTimeout(() => {
                if (!excelDataCache || excelDataCache.length === 0) {
                    return;
                }
                currentFilterConditions.searchTerm = searchTerm;
                currentPage = 1;
                applyFiltersAndRender(1);
            }, 300);
        });
    }
    
    // 초기화 버튼
    const resetBtn = document.getElementById('excel-reset-btn');
    if (resetBtn) {
        resetBtn.addEventListener('click', () => {
            if (searchInput) {
                searchInput.value = '';
            }
            currentFilterConditions.searchTerm = '';
            currentPage = 1;
            if (currentSheetName && excelDataCache) {
                applyFiltersAndRender(1);
            }
        });
    }
}

// 관리자 화면 탭 클릭 시 엑셀 데이터 로드는 index.html에서 처리됨

// 컬럼 리사이즈 기능 초기화
function initializeColumnResize() {
    const tables = ['settled-detail-table', 'unsettled-detail-table'];
    
    tables.forEach(tableId => {
        const table = document.getElementById(tableId);
        if (!table) return;
        
        // 정산금액 컬럼 (5번째)과 비고 컬럼 (6번째)에 리사이즈 핸들 추가
        const amountHeader = table.querySelector('thead th:nth-child(5)');
        const noteHeader = table.querySelector('thead th:nth-child(6)');
        
        if (amountHeader) {
            addResizeHandle(amountHeader, tableId, 5);
        }
        if (noteHeader) {
            addResizeHandle(noteHeader, tableId, 6);
        }
    });
    
    // 월별 정산 요약 테이블의 정산월 열(1번째)과 정산금액 열(2번째)에 리사이즈 핸들 추가
    const monthlyTable = document.getElementById('monthly-summary-table');
    if (monthlyTable) {
        const monthHeader = monthlyTable.querySelector('thead th:nth-child(1)');
        const amountHeader = monthlyTable.querySelector('thead th:nth-child(2)');
        if (monthHeader) {
            addResizeHandle(monthHeader, 'monthly-summary-table', 1);
        }
        if (amountHeader) {
            addResizeHandle(amountHeader, 'monthly-summary-table', 2);
        }
    }
}

// 관리자 테이블 리사이즈 기능 초기화
function initializeAdminTableResize(columnCount) {
    const adminTable = document.querySelector('.admin-table');
    if (!adminTable) return;
    
    // 모든 헤더 열에 리사이즈 핸들 추가
    for (let i = 1; i <= columnCount; i++) {
        const header = adminTable.querySelector(`thead th:nth-child(${i})`);
        if (header) {
            addResizeHandle(header, 'admin-table', i, true); // 관리자 테이블은 클래스 선택자 사용
        }
    }
}

// 리사이즈 핸들 추가
function addResizeHandle(header, tableId, columnIndex, useClassSelector = false) {
    // 기존 핸들 제거
    const existingHandle = header.querySelector('.resize-handle');
    if (existingHandle) {
        existingHandle.remove();
    }
    
    const handle = document.createElement('div');
    handle.className = 'resize-handle';
    header.style.position = 'relative';
    header.appendChild(handle);
    
    let isResizing = false;
    let startX = 0;
    let startWidth = 0;
    
    handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        
        isResizing = true;
        startX = e.pageX;
        startWidth = header.offsetWidth;
        
        const table = useClassSelector 
            ? document.querySelector(`.${tableId}`)
            : document.getElementById(tableId);
        if (table) {
            table.classList.add('resizing');
        }
        handle.classList.add('active');
        
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
    });
    
    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        
        e.preventDefault();
        const diff = e.pageX - startX;
        const newWidth = Math.max(50, startWidth + diff); // 최소 너비 50px
        
        // 컬럼 너비 설정
        const styleId = `${tableId}-col-${columnIndex}-style`;
        let style = document.getElementById(styleId);
        if (!style) {
            style = document.createElement('style');
            style.id = styleId;
            document.head.appendChild(style);
        }
        
        // 선택자 결정 (ID 또는 클래스)
        const selector = useClassSelector ? `.${tableId}` : `#${tableId}`;
        
        // 합계 셀 ID 확인
        let totalCellId = '';
        if (columnIndex === 5) {
            // 상세 내역 테이블의 정산금액 열
            totalCellId = tableId === 'settled-detail-table' ? '#totalAmountCell' : '#total-unsettled-detail';
        } else if (columnIndex === 2 && tableId === 'monthly-summary-table') {
            // 월별 정산 요약 테이블의 정산금액 열
            totalCellId = '#total-settled';
        }
        
        style.textContent = `
            ${selector} th:nth-child(${columnIndex}),
            ${selector} tbody td:nth-child(${columnIndex}),
            ${selector} tfoot td:nth-child(${columnIndex})${totalCellId ? `,
            ${totalCellId}` : ''} {
                width: ${newWidth}px !important;
                min-width: ${newWidth}px;
            }
        `;
    });
    
    document.addEventListener('mouseup', () => {
        if (!isResizing) return;
        
        isResizing = false;
        const table = useClassSelector 
            ? document.querySelector(`.${tableId}`)
            : document.getElementById(tableId);
        if (table) {
            table.classList.remove('resizing');
        }
        handle.classList.remove('active');
        
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
    });
}

// 테이블 업데이트 후 리사이즈 핸들 재설정
function setupResizeHandlesAfterUpdate() {
    // 약간의 지연 후 리사이즈 핸들 재설정 (DOM 업데이트 완료 후)
    setTimeout(() => {
        initializeColumnResize();
    }, 100);
}

// 정렬 헤더 클릭 처리
function handleSortClick(tableType, column) {
    const state = sortState[tableType];
    
    // 같은 컬럼 클릭 시 방향 토글, 다른 컬럼 클릭 시 오름차순으로 시작
    if (state.column === column) {
        state.direction = state.direction === 'asc' ? 'desc' : 'asc';
    } else {
        state.column = column;
        state.direction = 'asc';
    }
    
    // 헤더 UI 업데이트
    updateSortHeaders(tableType);
    
    // 테이블 재렌더링
    if (tableType === 'monthly') {
        // 🔥 월별 정산 요약: 항상 현재 상세 내역에서 계산 (함수 내부에서 처리)
        // originalMonthlyData는 정렬을 위해 저장된 것이지만, 재계산을 위해 함수 내부에서 처리
        updateMonthlySummary();
    } else if (tableType === 'settled') {
        updateSettledDetail(originalSettledDetail);
    } else if (tableType === 'unsettled') {
        updateUnsettledDetail(originalUnsettledDetail);
    }
}

// 정렬 헤더 UI 업데이트
function updateSortHeaders(tableType) {
    let tableId;
    
    if (tableType === 'monthly') {
        tableId = 'monthly-summary-table';
    } else if (tableType === 'settled') {
        tableId = 'settled-detail-table';
    } else if (tableType === 'unsettled') {
        tableId = 'unsettled-detail-table';
    } else {
        return;
    }
    
    const table = document.getElementById(tableId);
    if (!table) return;
    
    const sortableHeaders = table.querySelectorAll('th.sortable');
    const state = sortState[tableType];
    
    sortableHeaders.forEach((th) => {
        const column = th.getAttribute('data-column');
        th.classList.remove('asc', 'desc');
        
        if (state.column && column === state.column) {
            th.classList.add(state.direction);
        }
    });
}

// 정렬 헤더 클릭 이벤트 리스너 설정
function setupSortHeaders() {
    // 월별 정산 요약 테이블
    const monthlyTable = document.getElementById('monthly-summary-table');
    if (monthlyTable) {
        const sortableHeaders = monthlyTable.querySelectorAll('th.sortable');
        sortableHeaders.forEach(th => {
            th.addEventListener('click', () => {
                const column = th.getAttribute('data-column');
                handleSortClick('monthly', column);
            });
        });
    }
    
    // 월 정산 상세 내역 테이블
    const settledTable = document.getElementById('settled-detail-table');
    if (settledTable) {
        const sortableHeaders = settledTable.querySelectorAll('th.sortable');
        sortableHeaders.forEach(th => {
            th.addEventListener('click', () => {
                const column = th.getAttribute('data-column');
                handleSortClick('settled', column);
            });
        });
    }
    
    // 미정산 상세 내역 테이블
    const unsettledTable = document.getElementById('unsettled-detail-table');
    if (unsettledTable) {
        const sortableHeaders = unsettledTable.querySelectorAll('th.sortable');
        sortableHeaders.forEach(th => {
            th.addEventListener('click', () => {
                const column = th.getAttribute('data-column');
                handleSortClick('unsettled', column);
            });
        });
    }
}

// 앱 초기화
initializeApp();
setupAdminExcelDataHandlers();

// 필터 상태 저장
let filterState = {
    settled: {
        month: null,
        paymentDate: null,
        merchant: null,
        accountName: null,
        amount: null,
        note: null
    }
};

// 필터 기능 초기화
let filterInitialized = false;
function initializeFilters() {
    const settledTable = document.getElementById('settled-detail-table');
    if (!settledTable) return;
    
    // 이벤트 위임 사용: 테이블에 한 번만 이벤트 리스너 등록
    if (!filterInitialized) {
        settledTable.addEventListener('click', (e) => {
            const filterIcon = e.target.closest('.filter-icon');
            if (filterIcon) {
                e.stopPropagation();
                const column = filterIcon.getAttribute('data-column');
                toggleFilterDropdown(column);
            }
        });
        
        // 외부 클릭 시 드롭다운 닫기 (한 번만 등록)
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.filter-dropdown') && !e.target.closest('.filter-icon')) {
                closeAllFilterDropdowns();
            }
        });
        
        filterInitialized = true;
        console.log('✅ 필터 기능 초기화 완료');
    }
}

// 필터 드롭다운 토글
function toggleFilterDropdown(column) {
    const settledTable = document.getElementById('settled-detail-table');
    if (!settledTable) {
        console.error('❌ settled-detail-table을 찾을 수 없습니다.');
        return;
    }
    
    console.log(`🔍 필터 드롭다운 토글: ${column}`);
    
    // 다른 드롭다운 닫기
    closeAllFilterDropdowns();
    
    // 현재 컬럼의 드롭다운 찾기 또는 생성
    let dropdown = document.getElementById(`filter-dropdown-${column}`);
    const isOpening = !dropdown || !dropdown.classList.contains('active');
    
    if (!dropdown) {
        console.log(`📋 필터 드롭다운 생성: ${column}`);
        dropdown = createFilterDropdown(column);
        const th = settledTable.querySelector(`th[data-column="${column}"]`);
        if (th) {
            th.appendChild(dropdown);
            console.log(`✅ 필터 드롭다운 추가 완료: ${column}`);
        } else {
            console.error(`❌ th[data-column="${column}"] 요소를 찾을 수 없습니다.`);
            // 모든 th 요소 확인
            const allThs = settledTable.querySelectorAll('th');
            console.log('📋 사용 가능한 th 요소들:');
            allThs.forEach((th, idx) => {
                console.log(`   ${idx + 1}. data-column="${th.getAttribute('data-column')}", 클래스="${th.className}"`);
            });
        }
    }
    
    // 드롭다운이 열릴 때 옵션 목록 업데이트 (최신 데이터 반영)
    const optionsDiv = dropdown.querySelector('.filter-options');
    if (optionsDiv && isOpening) {
        populateFilterOptions(column, optionsDiv);
        // 검색 입력 초기화
        const searchInput = dropdown.querySelector('.filter-search input');
        if (searchInput) {
            searchInput.value = '';
        }
    }
    
    // 드롭다운 표시/숨김
    dropdown.classList.toggle('active');
    
    // 필터 아이콘 활성화 상태 업데이트
    const filterIcon = settledTable.querySelector(`.filter-icon[data-column="${column}"]`);
    if (filterIcon) {
        if (dropdown.classList.contains('active')) {
            filterIcon.classList.add('active');
        } else {
            updateFilterIconState(filterIcon, column);
        }
    }
}

// 모든 필터 드롭다운 닫기
function closeAllFilterDropdowns() {
    const dropdowns = document.querySelectorAll('.filter-dropdown');
    dropdowns.forEach(dropdown => {
        dropdown.classList.remove('active');
    });
    
    // 모든 필터 아이콘 상태 업데이트
    const filterIcons = document.querySelectorAll('.filter-icon');
    filterIcons.forEach(icon => {
        const column = icon.getAttribute('data-column');
        updateFilterIconState(icon, column);
    });
}

// 필터 아이콘 상태 업데이트
function updateFilterIconState(icon, column) {
    const hasFilter = filterState.settled[column] !== null && 
                      filterState.settled[column].length > 0;
    if (hasFilter) {
        icon.classList.add('active');
    } else {
        icon.classList.remove('active');
    }
}

// 필터 드롭다운 생성
function createFilterDropdown(column) {
    const dropdown = document.createElement('div');
    dropdown.id = `filter-dropdown-${column}`;
    dropdown.className = 'filter-dropdown';
    
    // 검색 입력
    const searchDiv = document.createElement('div');
    searchDiv.className = 'filter-search';
    const searchInput = document.createElement('input');
    searchInput.type = 'text';
    searchInput.placeholder = '(모두)에서 검색';
    searchInput.addEventListener('input', (e) => {
        filterOptions(dropdown, e.target.value);
    });
    searchDiv.appendChild(searchInput);
    
    // 옵션 목록
    const optionsDiv = document.createElement('div');
    optionsDiv.className = 'filter-options';
    optionsDiv.id = `filter-options-${column}`;
    
    // 액션 버튼 (엑셀 스타일: 확인/취소만)
    const actionsDiv = document.createElement('div');
    actionsDiv.className = 'filter-actions';
    
    const cancelBtn = document.createElement('button');
    cancelBtn.className = 'btn-cancel';
    cancelBtn.textContent = '취소';
    cancelBtn.addEventListener('click', () => {
        // 변경사항 취소하고 드롭다운 닫기
        closeAllFilterDropdowns();
    });
    
    const confirmBtn = document.createElement('button');
    confirmBtn.className = 'btn-confirm';
    confirmBtn.textContent = '확인';
    confirmBtn.addEventListener('click', () => {
        applyFilter(column);
        closeAllFilterDropdowns();
    });
    
    actionsDiv.appendChild(cancelBtn);
    actionsDiv.appendChild(confirmBtn);
    
    dropdown.appendChild(searchDiv);
    dropdown.appendChild(optionsDiv);
    dropdown.appendChild(actionsDiv);
    
    // 옵션 목록 생성
    populateFilterOptions(column, optionsDiv);
    
    return dropdown;
}

// 필터 옵션 목록 생성
function populateFilterOptions(column, optionsDiv) {
    const data = originalSettledDetail || [];
    console.log(`📋 populateFilterOptions 호출: column=${column}, data.length=${data.length}`);
    
    if (data.length === 0) {
        optionsDiv.innerHTML = '<div style="padding: 8px; color: #999;">데이터가 없습니다.</div>';
        console.warn(`⚠️ 필터 옵션 생성 실패: 데이터가 없습니다. (column: ${column})`);
        return;
    }
    
    // 날짜 컬럼인지 확인
    const isDateColumn = column === 'month' || column === 'paymentDate';
    console.log(`📋 날짜 컬럼 여부: ${isDateColumn} (column: ${column})`);
    
    // 현재 필터 상태 가져오기
    const currentFilter = filterState.settled[column];
    
    optionsDiv.innerHTML = '';
    
    if (isDateColumn) {
        // 날짜 컬럼: 연도/월 계층 구조로 표시
        const dateMap = new Map(); // 연도별 월 목록
        
        data.forEach(item => {
            let value = '';
            if (column === 'month') {
                value = item.settlementMonth || item.month || '';
            } else if (column === 'paymentDate') {
                value = item.paymentDate || '';
            }
            
            if (value) {
                // YYYY-MM 형식에서 연도와 월 추출
                const match = value.match(/^(\d{4})-(\d{2})/);
                if (match) {
                    const year = match[1];
                    const month = match[2];
                    if (!dateMap.has(year)) {
                        dateMap.set(year, new Set());
                    }
                    dateMap.get(year).add(month);
                } else {
                    // YYYY-MM-DD 형식
                    const match2 = value.match(/^(\d{4})-(\d{2})-\d{2}/);
                    if (match2) {
                        const year = match2[1];
                        const month = match2[2];
                        if (!dateMap.has(year)) {
                            dateMap.set(year, new Set());
                        }
                        dateMap.get(year).add(month);
                    }
                }
            }
        });
        
        // "(모두 선택)" 체크박스 추가
        const allValues = [];
        dateMap.forEach((months, year) => {
            months.forEach(month => {
                allValues.push(`${year}-${month}`);
            });
        });
        const selectedValues = (currentFilter && currentFilter.length > 0) ? currentFilter : allValues;
        const allSelected = allValues.every(v => selectedValues.includes(v));
        
        const selectAllDiv = createFilterOption('(모두 선택)', '', true, allSelected, column, () => {
            toggleSelectAll(column, optionsDiv, !allSelected);
        });
        optionsDiv.appendChild(selectAllDiv);
        
        // 연도별로 정렬
        const sortedYears = Array.from(dateMap.keys()).sort();
        
        sortedYears.forEach(year => {
            const months = Array.from(dateMap.get(year)).sort();
            
            // 연도 헤더
            const yearDiv = createFilterOption(`${year}년`, year, true, false, column, () => {
                toggleYearSelection(column, year, months, optionsDiv);
            }, true);
            yearDiv.dataset.year = year;
            yearDiv.dataset.expanded = 'true';
            optionsDiv.appendChild(yearDiv);
            
            // 월 옵션들
            months.forEach(month => {
                const monthValue = `${year}-${month}`;
                const monthLabel = `${month}월`;
                const isSelected = selectedValues.includes(monthValue);
                const monthDiv = createFilterOption(monthLabel, monthValue, false, isSelected, column, null, false, true);
                monthDiv.style.display = 'block'; // 기본적으로 표시
                optionsDiv.appendChild(monthDiv);
            });
        });
    } else {
        // 일반 컬럼: 단순 리스트
        const uniqueValues = new Set();
        data.forEach(item => {
            let value = '';
            switch(column) {
                case 'merchant':
                    value = item.merchant || '';
                    break;
                case 'accountName':
                    value = item.accountName || '';
                    break;
                case 'amount':
                    value = formatNumber(item.amount || 0);
                    break;
                case 'note':
                    value = item.note || '';
                    break;
            }
            if (value !== '') {
                uniqueValues.add(String(value));
            }
        });
        
        // 정렬된 값 목록
        const sortedValues = Array.from(uniqueValues).sort((a, b) => {
            // 숫자 형식인 경우 숫자로 정렬
            const numA = parseFloat(a.replace(/,/g, ''));
            const numB = parseFloat(b.replace(/,/g, ''));
            if (!isNaN(numA) && !isNaN(numB)) {
                return numA - numB;
            }
            return a.localeCompare(b, 'ko');
        });
        
        const selectedValues = (currentFilter && currentFilter.length > 0) ? currentFilter : sortedValues;
        const allSelected = sortedValues.every(v => selectedValues.includes(v));
        
        // "(모두 선택)" 체크박스 추가
        const selectAllDiv = createFilterOption('(모두 선택)', '', true, allSelected, column, () => {
            toggleSelectAll(column, optionsDiv, !allSelected);
        });
        optionsDiv.appendChild(selectAllDiv);
        
        // 옵션 생성
        sortedValues.forEach(value => {
            const isSelected = selectedValues.includes(value);
            const optionDiv = createFilterOption(value, value, false, isSelected, column);
            optionsDiv.appendChild(optionDiv);
        });
    }
}

// 필터 옵션 생성 헬퍼 함수
function createFilterOption(label, value, isSelectAll, isChecked, column, onClick = null, isParent = false, isChild = false) {
    const optionDiv = document.createElement('div');
    optionDiv.className = 'filter-option';
    if (isParent) {
        optionDiv.classList.add('parent');
    }
    if (isChild) {
        optionDiv.classList.add('child');
    }
    
    // 확장/축소 아이콘 (부모인 경우)
    if (isParent) {
        const expandIcon = document.createElement('span');
        expandIcon.className = 'filter-expand-icon';
        expandIcon.textContent = '−'; // 마이너스 (확장됨)
        expandIcon.addEventListener('click', (e) => {
            e.stopPropagation();
            toggleYearExpand(optionDiv);
        });
        optionDiv.appendChild(expandIcon);
    } else if (isChild) {
        const expandIcon = document.createElement('span');
        expandIcon.className = 'filter-expand-icon';
        expandIcon.textContent = '+';
        expandIcon.style.visibility = 'hidden'; // 자식은 아이콘 숨김
        optionDiv.appendChild(expandIcon);
    }
    
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.id = `filter-${column}-${value || label}`;
    checkbox.value = value || label;
    checkbox.checked = isChecked;
    checkbox.addEventListener('change', () => {
        if (onClick) {
            onClick();
        } else {
            updateFilterState(column);
            updateSelectAllState(column);
        }
    });
    
    const labelEl = document.createElement('label');
    labelEl.htmlFor = `filter-${column}-${value || label}`;
    labelEl.textContent = label;
    
    optionDiv.appendChild(checkbox);
    optionDiv.appendChild(labelEl);
    
    return optionDiv;
}

// 전체 선택/해제
function toggleSelectAll(column, optionsDiv, selectAll) {
    const checkboxes = optionsDiv.querySelectorAll('input[type="checkbox"]:not([value=""])');
    checkboxes.forEach(checkbox => {
        if (checkbox.closest('.filter-option').style.display !== 'none') {
            checkbox.checked = selectAll;
        }
    });
    updateFilterState(column);
}

// 전체 선택 상태 업데이트
function updateSelectAllState(column) {
    const dropdown = document.getElementById(`filter-dropdown-${column}`);
    if (!dropdown) return;
    
    const optionsDiv = dropdown.querySelector('.filter-options');
    const allCheckbox = optionsDiv.querySelector('input[type="checkbox"][value=""]');
    if (!allCheckbox) return;
    
    const allCheckboxes = optionsDiv.querySelectorAll('input[type="checkbox"]:not([value=""])');
    const checkedCount = Array.from(allCheckboxes).filter(cb => cb.checked && cb.closest('.filter-option').style.display !== 'none').length;
    const visibleCount = Array.from(allCheckboxes).filter(cb => cb.closest('.filter-option').style.display !== 'none').length;
    
    allCheckbox.checked = checkedCount === visibleCount && visibleCount > 0;
}

// 연도 선택 토글
function toggleYearSelection(column, year, months, optionsDiv) {
    const yearDiv = optionsDiv.querySelector(`[data-year="${year}"]`);
    if (!yearDiv) return;
    
    const yearCheckbox = yearDiv.querySelector('input[type="checkbox"]');
    const isChecked = yearCheckbox.checked;
    
    // 해당 연도의 모든 월 체크박스 업데이트
    months.forEach(month => {
        const monthValue = `${year}-${month}`;
        const monthCheckbox = optionsDiv.querySelector(`input[type="checkbox"][value="${monthValue}"]`);
        if (monthCheckbox && monthCheckbox.closest('.filter-option').style.display !== 'none') {
            monthCheckbox.checked = isChecked;
        }
    });
    
    updateFilterState(column);
    updateSelectAllState(column);
}

// 연도 확장/축소 토글
function toggleYearExpand(yearDiv) {
    const year = yearDiv.dataset.year;
    const isExpanded = yearDiv.dataset.expanded === 'true';
    const expandIcon = yearDiv.querySelector('.filter-expand-icon');
    
    const optionsDiv = yearDiv.parentElement;
    const allOptions = Array.from(optionsDiv.querySelectorAll('.filter-option'));
    const yearIndex = allOptions.indexOf(yearDiv);
    
    // 다음 연도까지의 모든 자식 옵션 찾기
    let endIndex = allOptions.length;
    for (let i = yearIndex + 1; i < allOptions.length; i++) {
        if (allOptions[i].classList.contains('parent')) {
            endIndex = i;
            break;
        }
    }
    
    // 자식 옵션 표시/숨김
    for (let i = yearIndex + 1; i < endIndex; i++) {
        if (allOptions[i].classList.contains('child')) {
            allOptions[i].style.display = isExpanded ? 'none' : 'flex';
        }
    }
    
    // 아이콘 업데이트
    expandIcon.textContent = isExpanded ? '+' : '−';
    yearDiv.dataset.expanded = isExpanded ? 'false' : 'true';
}

// 필터 옵션 검색
function filterOptions(dropdown, searchText) {
    const optionsDiv = dropdown.querySelector('.filter-options');
    const options = optionsDiv.querySelectorAll('.filter-option');
    
    if (!searchText || searchText.trim() === '') {
        // 검색어가 없으면 모두 표시 (계층 구조 유지)
        options.forEach(option => {
            option.style.display = 'flex';
        });
        return;
    }
    
    const searchLower = searchText.toLowerCase();
    options.forEach(option => {
        const label = option.querySelector('label');
        const text = label.textContent.toLowerCase();
        const isSelectAll = label.textContent === '(모두 선택)';
        
        // "(모두 선택)"은 항상 표시
        if (isSelectAll) {
            option.style.display = 'flex';
            return;
        }
        
        // 검색어와 일치하면 표시
        if (text.includes(searchLower)) {
            option.style.display = 'flex';
            // 자식인 경우 부모 연도도 표시하고 확장
            if (option.classList.contains('child')) {
                const yearDiv = findParentYear(option);
                if (yearDiv) {
                    yearDiv.style.display = 'flex';
                    // 연도가 축소되어 있으면 확장
                    if (yearDiv.dataset.expanded === 'false') {
                        toggleYearExpand(yearDiv);
                    }
                }
            }
        } else {
            option.style.display = 'none';
        }
    });
}

// 자식 옵션의 부모 연도 찾기
function findParentYear(childOption) {
    let current = childOption.previousElementSibling;
    while (current) {
        if (current.classList.contains('parent')) {
            return current;
        }
        current = current.previousElementSibling;
    }
    return null;
}

// 전체 선택/해제
function selectAllOptions(dropdown, column, selectAll) {
    const optionsDiv = dropdown.querySelector('.filter-options');
    const checkboxes = optionsDiv.querySelectorAll('input[type="checkbox"]');
    
    checkboxes.forEach(checkbox => {
        if (checkbox.closest('.filter-option').style.display !== 'none') {
            checkbox.checked = selectAll;
        }
    });
    
    updateFilterState(column);
}

// 필터 상태 업데이트
function updateFilterState(column) {
    const dropdown = document.getElementById(`filter-dropdown-${column}`);
    if (!dropdown) return;
    
    const optionsDiv = dropdown.querySelector('.filter-options');
    const checkboxes = optionsDiv.querySelectorAll('input[type="checkbox"]:checked');
    const selectedValues = Array.from(checkboxes)
        .filter(cb => cb.value !== '') // "(모두 선택)" 체크박스 제외
        .map(cb => cb.value);
    
    // 모든 값이 선택되었거나 아무것도 선택되지 않았으면 null로 설정 (필터 없음)
    const allCheckboxes = optionsDiv.querySelectorAll('input[type="checkbox"]:not([value=""])');
    const allValues = Array.from(allCheckboxes).map(cb => cb.value);
    const allSelected = allValues.length > 0 && allValues.every(v => selectedValues.includes(v));
    
    if (allSelected || selectedValues.length === 0) {
        filterState.settled[column] = null;
    } else {
        filterState.settled[column] = selectedValues;
    }
    
    console.log(`📊 필터 상태 업데이트 [${column}]:`, filterState.settled[column]);
}

// 필터 적용
function applyFilter(column) {
    updateFilterState(column);
    applyAllFilters();
}

// 데이터에 필터 적용 (내부 함수)
function applyFiltersToData(data) {
    let filteredData = [...data];
    
    // 각 컬럼별 필터 적용
    Object.keys(filterState.settled).forEach(col => {
        const filter = filterState.settled[col];
        if (filter && filter.length > 0) {
            const beforeCount = filteredData.length;
            filteredData = filteredData.filter(item => {
                let value = '';
                switch(col) {
                    case 'month':
                        value = item.settlementMonth || item.month || '';
                        break;
                    case 'paymentDate':
                        value = item.paymentDate || '';
                        // YYYY-MM-DD 형식에서 YYYY-MM 추출
                        if (value && value.match(/^\d{4}-\d{2}-\d{2}$/)) {
                            value = value.substring(0, 7);
                        }
                        break;
                    case 'merchant':
                        value = item.merchant || '';
                        break;
                    case 'accountName':
                        value = item.accountName || '';
                        break;
                    case 'amount':
                        value = formatNumber(item.amount || 0);
                        break;
                    case 'note':
                        value = item.note || '';
                        break;
                }
                const stringValue = String(value);
                // 날짜 컬럼의 경우 YYYY-MM 형식으로 비교
                if (col === 'month' || col === 'paymentDate') {
                    return filter.some(f => {
                        // 필터 값이 YYYY-MM 형식이면 정확히 일치
                        if (f.match(/^\d{4}-\d{2}$/)) {
                            return stringValue.startsWith(f);
                        }
                        return stringValue === f;
                    });
                }
                return filter.includes(stringValue);
            });
            const afterCount = filteredData.length;
            if (beforeCount !== afterCount) {
                console.log(`🔍 필터 적용 [${col}]: ${beforeCount}개 → ${afterCount}개`);
            }
        }
    });
    
    return filteredData;
}

// 모든 필터 적용
function applyAllFilters() {
    // 원본 데이터에 필터를 적용하여 테이블 업데이트
    // skipOriginalSave를 true로 설정하여 원본 데이터를 덮어쓰지 않음
    updateSettledDetail(originalSettledDetail, true);
    
    // 필터 아이콘 상태 업데이트
    const settledTable = document.getElementById('settled-detail-table');
    if (settledTable) {
        const filterIcons = settledTable.querySelectorAll('.filter-icon');
        filterIcons.forEach(icon => {
            const col = icon.getAttribute('data-column');
            updateFilterIconState(icon, col);
        });
    }
    
    console.log('✅ 필터 적용 완료:', {
        원본데이터개수: originalSettledDetail.length,
        필터상태: filterState.settled
    });
}

// 컬럼 리사이즈 기능 초기화
document.addEventListener('DOMContentLoaded', () => {
    initializeColumnResize();
    setupSortHeaders();
    initializeFilters();
});

// 테스트용 데이터 표시 제거 (프로덕션에서는 불필요)
// fetch("http://localhost:3000/api/all-data")
//   .then(res => res.json())
//   .then(data => {
//     console.log("받은 데이터:", data);
//     const testElement = document.getElementById("test");
//     if (testElement) {
//       testElement.innerText = JSON.stringify(data, null, 2);
//     }
//   });


