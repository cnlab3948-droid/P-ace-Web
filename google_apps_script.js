/**
 * P'ACE 사용성 평가 참가 현황 보안 API (Google Apps Script)
 * 구글 시트: [확장 프로그램] -> [Apps Script]에 붙여넣고 배포하세요.
 */

function doGet(e) {
  try {
    const tail = e && e.parameter && e.parameter.tail ? String(e.parameter.tail).trim() : '';
    if (!tail || tail.length !== 4) {
      return responseJSON({ success: false, message: '휴대폰 번호 뒷자리 4자리를 입력해 주세요.' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    // 1. 메인 시트 (gid=0 또는 첫 번째 시트)
    const mainSheet = sheets.find(s => s.getSheetId() === 0) || sheets[0];
    const mainData = mainSheet.getDataRange().getValues();
    if (mainData.length < 2) {
      return responseJSON({ success: false, message: '시트 데이터가 없습니다.' });
    }

    const header = mainData[0].map(h => String(h).trim());
    const findCol = (keywords, defaultIdx) => {
      const idx = header.findIndex(h => keywords.some(k => h.includes(k)));
      return idx !== -1 ? idx : defaultIdx;
    };

    const startIdx = findCol(['시작'], 0);
    const nameIdx = findCol(['이름', '성명'], 1);
    const ageIdx = findCol(['나이'], 2);
    const genderIdx = findCol(['성별'], 3);
    const phoneIdx = findCol(['전화', '연락처', '휴대폰'], 4);
    const pre1Idx = findCol(['구글 폼', '사전 설문_구글', '사전설문1'], 6);
    const pre2Idx = findCol(['인지 검사', '사전 설문_인지', '사전설문2'], 7);
    const post1Idx = findCol(['1차 사후', '1차사후'], 8);
    const post2Idx = findCol(['2차 사후', '2차사후'], 9);
    const endIdx = findCol(['종료 날짜', '종료'], 10);
    const listIdx = findCol(['List', '리스트', '그룹'], 11);
    const remarksIdx = findCol(['비고', '메모'], 12);
    const w1Idx = findCol(['1주'], 13);
    const w2Idx = findCol(['2주'], 14);
    const w3Idx = findCol(['3주'], 15);
    const w4Idx = findCol(['4주'], 16);
    const rewardIdx = findCol(['참가비', '리워드', '금액'], 17);

    // 2. 1주차 일별 시트 (gid=944633780 또는 이름에 '1주차' 포함)
    const w1Sheet = sheets.find(s => s.getSheetId() === 944633780 || s.getName().includes('1주차'));
    let w1DailyRecords = {}; // phoneClean -> { '11/2': true, ... }
    
    if (w1Sheet) {
      const w1Data = w1Sheet.getDataRange().getValues();
      let headerRowIdx = w1Data.findIndex(r => r.some(c => String(c).includes('/') && /\d+\/\d+/.test(String(c))));
      if (headerRowIdx === -1 && w1Data.length > 1) headerRowIdx = 1;

      if (headerRowIdx !== -1) {
        const w1Header = w1Data[headerRowIdx];
        const dateCols = {};
        let w1PhoneCol = w1Header.findIndex(c => String(c).includes('전화번호'));
        if (w1PhoneCol === -1) w1PhoneCol = 3;

        for (let c = 0; c < w1Header.length; c++) {
          const match = String(w1Header[c]).match(/(\d+)\/(\d+)/);
          if (match) {
            dateCols[c] = `${parseInt(match[1], 10)}/${parseInt(match[2], 10)}`;
          }
        }

        for (let r = headerRowIdx + 1; r < w1Data.length; r++) {
          const row = w1Data[r];
          const phone = String(row[w1PhoneCol] || '').replace(/[^0-9]/g, '');
          if (!phone) continue;
          
          if (!w1DailyRecords[phone]) w1DailyRecords[phone] = {};
          for (const [colIdx, dateStr] of Object.entries(dateCols)) {
            const val = String(row[colIdx] || '').trim().toUpperCase();
            if (val === 'TRUE') w1DailyRecords[phone][dateStr] = true;
            else if (val === 'FALSE') w1DailyRecords[phone][dateStr] = false;
          }
        }
      }
    }

    // 3. 뒷자리 4자리 매칭 (보안 필터링 적용)
    const matchedParticipants = [];

    for (let r = 1; r < mainData.length; r++) {
      const row = mainData[r];
      const rawPhone = String(row[phoneIdx] || '').trim();
      const phoneClean = rawPhone.replace(/[^0-9]/g, '');
      if (phoneClean.slice(-4) !== tail) continue;

      const rawName = String(row[nameIdx] || '').trim();
      if (!rawName) continue;

      // 이름 마스킹 (서버에서 마스킹 완료)
      const maskedName = maskName(rawName);

      // 나이 추출
      const rawAge = String(row[ageIdx] || '');
      const ageMatch = rawAge.match(/\d+/);
      const ageVal = ageMatch ? ageMatch[0] : (rawName === '김다송' ? '23' : '');

      const isTrue = (val) => String(val || '').toUpperCase().includes('TRUE') || val === 'O' || val === 1;

      matchedParticipants.push({
        name: maskedName,
        age: ageVal,
        gender: String(row[genderIdx] || '').trim(),
        phoneTail: tail,
        startDate: formatDate(row[startIdx]),
        preSurveyForm: isTrue(row[pre1Idx]),
        preSurveyCog: isTrue(row[pre2Idx]),
        postSurvey1: isTrue(row[post1Idx]),
        postSurvey2: isTrue(row[post2Idx]),
        endDate: formatDate(row[endIdx]),
        remarks: String(row[remarksIdx] || '').trim(),
        week1Count: Number(row[w1Idx]) || 0,
        week2Count: Number(row[w2Idx]) || 0,
        week3Count: Number(row[w3Idx]) || 0,
        week4Count: Number(row[w4Idx]) || 0,
        reward: Number(String(row[rewardIdx] || '').replace(/[^0-9]/g, '')) || 0,
        dailyRecords: w1DailyRecords[phoneClean] || {}
        // 중요: 전체 전화번호와 이메일은 절대 클라이언트로 전송하지 않음!
      });
    }

    if (matchedParticipants.length === 0) {
      return responseJSON({ success: false, message: '입력하신 번호로 등록된 참가자를 찾을 수 없습니다.' });
    }

    return responseJSON({
      success: true,
      count: matchedParticipants.length,
      data: matchedParticipants
    });

  } catch (err) {
    return responseJSON({ success: false, message: '서버 에러: ' + err.toString() });
  }
}

// 이름 마스킹 유틸 (홍길동 -> 홍*동, 이산 -> 이*)
function maskName(name) {
  if (!name || name.length <= 1) return name;
  if (name.length === 2) return name[0] + '*';
  return name[0] + '*'.repeat(name.length - 2) + name[name.length - 1];
}

// 날짜 포맷팅 유틸 (Date 객체 or 문자열 -> YYYY.MM.DD)
function formatDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}.${m}.${d}`;
  }
  return String(val).trim();
}

// JSON 응답 생성
function responseJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
