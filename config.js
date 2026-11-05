// 로아 레이드 일정 설정 — 메모장으로 열어서 따옴표 안만 바꾸세요. (index.html 과 함께 GitHub 에 올리기)
window.RAID_CONFIG = {
  // 공대원 입력 공유: 구글 Apps Script 웹 앱 주소 (https://script.google.com/macros/s/…/exec)
  // 비워 두면 '로컬 모드' — 입력한 시간이 각자 브라우저에만 저장돼요.
  teamSyncUrl: "",

  // 일정 만들 때 고르는 레이드 목록 (위에서부터 보여요). levels = 난이도 버튼
  // 새 레이드가 나오면 한 줄 추가, 끝난 레이드는 그 줄을 지우면 돼요.
  raids: [
    { name: "벨가르딘", tag: "그림자 레이드", levels: ["노말", "하드", "나이트메어"] },
    { name: "종막 익스트림", tag: "기간 한정 · 11/18까지", levels: ["노말", "하드", "나이트메어"] },
    { name: "종막", tag: "카제로스", levels: ["노말", "하드"] },
    { name: "4막", tag: "아르모체", levels: ["노말", "하드"] },
    { name: "3막", tag: "모르둠", levels: ["노말", "하드"] },
    { name: "2막", tag: "아브렐슈드", levels: ["노말", "하드"] },
    { name: "1막", tag: "에기르", levels: ["노말", "하드"] },
  ],
};
