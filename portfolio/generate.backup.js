const pptxgen = require('pptxgenjs');

// ── Pet Friendly Soft 컬러 팔레트 (코드용 — # 없이) ──────────────────────
const C = {
  bg: 'FBF7F1',          // Peach Cream (도미넌트)
  card: 'FFFFFF',
  fg: '1F1916',          // Warm Black
  muted: '6B5757',       // Warm Gray
  line: 'EDE4D6',
  accent: 'E07856',      // Soft Coral (메인 액센트)
  accentSoft: 'FFE3D6',
  sage: '88A876',        // Sage Green (보조 액센트)
  sageSoft: 'E5EFD7',
  gold: 'D4A24A',
  destructive: 'D63B3B',
  darkBg: '2A1F1F',      // 표지·마감용 다크
};

const F = {
  header: 'Georgia',     // Editorial 헤더
  body: 'Calibri',       // 본문
  korean: 'Malgun Gothic', // 한글 fallback (시스템에 따라 자동)
};

// ── 페이지 번호 (전역 상태) ───────────────────────────────────────────
const TOTAL = 16;
let currentPage = 0;

// ── 슬라이드 공통 헤더/푸터 ──────────────────────────────────────────
function drawShell(slide, label, opts = {}) {
  currentPage += 1;
  const onDark = opts.onDark === true;
  const fgColor = onDark ? 'FFFFFF' : C.fg;
  const mutedColor = onDark ? 'FFFFFF' : C.muted;

  // 좌상단 — 코랄 dot + label
  slide.addShape('ellipse', {
    x: 0.6, y: 0.42, w: 0.12, h: 0.12,
    fill: { color: C.accent }, line: { color: C.accent, width: 0 },
  });
  slide.addText(label, {
    x: 0.85, y: 0.36, w: 5, h: 0.25,
    fontFace: F.body, fontSize: 9, color: mutedColor, bold: true,
    charSpacing: 4, margin: 0,
  });

  // 우상단 — 페이지 번호
  const pageStr = String(currentPage).padStart(2, '0');
  slide.addText(`${pageStr} / ${String(TOTAL).padStart(2, '0')}`, {
    x: 11.8, y: 0.36, w: 1.2, h: 0.25,
    fontFace: F.body, fontSize: 9, color: mutedColor, bold: true,
    align: 'right', charSpacing: 3, margin: 0,
  });

  // 하단 푸터
  slide.addText('PAWMART · PORTFOLIO 2026', {
    x: 0.6, y: 7.2, w: 6, h: 0.2,
    fontFace: F.body, fontSize: 8, color: mutedColor,
    charSpacing: 4, margin: 0,
  });
  slide.addText('서인석 · seok2@pawmart.kr', {
    x: 7.3, y: 7.2, w: 5.4, h: 0.2,
    fontFace: F.body, fontSize: 8, color: mutedColor,
    align: 'right', margin: 0,
  });
}

// ── 카드 helper ──────────────────────────────────────────────
function drawCard(slide, x, y, w, h, opts = {}) {
  const fill = opts.fill ?? C.card;
  const border = opts.border ?? C.line;
  slide.addShape('roundRect', {
    x, y, w, h,
    fill: { color: fill },
    line: { color: border, width: 0.5 },
    rectRadius: 0.12,
  });
}

// 작은 액센트 막대 (왼쪽 4px coral)
function drawAccentBar(slide, x, y, h, color = C.accent) {
  slide.addShape('rect', {
    x, y, w: 0.05, h,
    fill: { color }, line: { color, width: 0 },
  });
}

// ── PPT 시작 ──────────────────────────────────────────────────────
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE'; // 13.333 x 7.5
pres.title = 'Pawmart Portfolio';
pres.author = '서인석';
pres.subject = '프리미엄 펫 e-commerce 포트폴리오';

// ════════════════════════════════════════════════════════════
// SLIDE 01 — 표지
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  // 우측 컬러 블록 (3등분)
  s.addShape('rect', { x: 9.0, y: 0, w: 4.4, h: 2.5, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addShape('rect', { x: 9.0, y: 2.5, w: 4.4, h: 2.5, fill: { color: C.accentSoft }, line: { color: C.accentSoft, width: 0 } });
  s.addShape('rect', { x: 9.0, y: 5.0, w: 4.4, h: 2.5, fill: { color: C.sage }, line: { color: C.sage, width: 0 } });

  // 좌상단 라벨
  s.addShape('ellipse', { x: 0.6, y: 0.42, w: 0.12, h: 0.12, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText('PORTFOLIO · 2026', {
    x: 0.85, y: 0.36, w: 5, h: 0.25,
    fontFace: F.body, fontSize: 10, color: 'FFFFFF', bold: true, charSpacing: 5, margin: 0,
  });

  // 메인 타이틀
  s.addText('PAWMART', {
    x: 0.6, y: 2.4, w: 8, h: 1.6,
    fontFace: F.header, fontSize: 96, bold: true, color: 'FFFFFF', charSpacing: -2, margin: 0,
  });

  // 부제
  s.addText('반려동물과 보호자를 위한 큐레이션 마켓', {
    x: 0.6, y: 4.05, w: 8, h: 0.4,
    fontFace: F.body, fontSize: 18, color: C.accentSoft, margin: 0,
  });

  // 한국어 부제 / Editorial italic
  s.addText('A Premium Pet E-Commerce Experience', {
    x: 0.6, y: 4.5, w: 8, h: 0.35,
    fontFace: F.header, italic: true, fontSize: 16, color: C.accent, margin: 0,
  });

  // 작성자
  s.addText('서인석 · Frontend & Backend', {
    x: 0.6, y: 6.6, w: 6, h: 0.3,
    fontFace: F.body, fontSize: 11, color: 'FFFFFF', charSpacing: 2, margin: 0,
  });
  s.addText('2026년 5월 · 풀스택 포트폴리오', {
    x: 0.6, y: 6.95, w: 6, h: 0.3,
    fontFace: F.body, fontSize: 9, color: 'FFFFFF', margin: 0,
  });
}
currentPage = 1; // 표지는 카운트 X, 다음 슬라이드부터 1

// ════════════════════════════════════════════════════════════
// SLIDE 02 — 왜 이 프로젝트인가
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'WHY THIS PROJECT');

  s.addText('왜 Pawmart를 만들었나', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('펫 시장은 빠르게 커지지만, 보호자가 겪는 불편은 그대로다.', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  // 좌측 — 큰 통계
  drawCard(s, 0.6, 2.7, 5.0, 4.0, { fill: C.accent });
  s.addText('6.2조 원', {
    x: 0.85, y: 3.3, w: 4.6, h: 1.5,
    fontFace: F.header, fontSize: 80, bold: true, color: 'FFFFFF', margin: 0,
  });
  s.addText('2026년 한국 펫 산업 규모', {
    x: 0.85, y: 4.85, w: 4.6, h: 0.4,
    fontFace: F.body, fontSize: 14, color: 'FFFFFF', charSpacing: 1, margin: 0,
  });
  s.addText('전년 대비 +9% · 출처: 농촌경제연구원 추정', {
    x: 0.85, y: 5.25, w: 4.6, h: 0.4,
    fontFace: F.body, fontSize: 10, color: 'FFFFFF', margin: 0,
  });

  // 우측 — 페인포인트 3개
  const pains = [
    {
      n: '01',
      t: '제품 정보가 흩어져 있다',
      d: '카탈로그·블로그·SNS 곳곳을 뒤져야 우리 아이에게 맞는 제품을 찾을 수 있다.',
    },
    {
      n: '02',
      t: '맞춤 추천이 부족하다',
      d: '품종·연령·체중·건강 상태에 따라 다른데, 일률적인 베스트셀러만 추천된다.',
    },
    {
      n: '03',
      t: '재구매가 번거롭다',
      d: '사료·모래는 매번 똑같이 검색해서 다시 산다. 정기배송 옵션이 없다.',
    },
  ];
  pains.forEach((p, i) => {
    const y = 2.7 + i * 1.35;
    drawCard(s, 6.0, y, 6.7, 1.2);
    drawAccentBar(s, 6.0, y, 1.2);
    s.addText(p.n, {
      x: 6.2, y: y + 0.15, w: 0.6, h: 0.35,
      fontFace: F.header, italic: true, fontSize: 22, color: C.accent, bold: true, margin: 0,
    });
    s.addText(p.t, {
      x: 6.85, y: y + 0.18, w: 5.7, h: 0.4,
      fontFace: F.body, fontSize: 15, bold: true, color: C.fg, margin: 0,
    });
    s.addText(p.d, {
      x: 6.85, y: y + 0.6, w: 5.7, h: 0.55,
      fontFace: F.body, fontSize: 11, color: C.muted, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 03 — 누구를 위한가 (페르소나)
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'TARGET USER');

  s.addText('누구를 위한 서비스인가', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('1차 타겟 페르소나 — 양육 5년차 직장인 보호자', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  // 좌측 페르소나 카드
  drawCard(s, 0.6, 2.7, 5.0, 4.0, { fill: C.card });
  s.addShape('ellipse', {
    x: 1.0, y: 3.0, w: 1.4, h: 1.4,
    fill: { color: C.accentSoft }, line: { color: C.accent, width: 2 },
  });
  s.addText('JY', {
    x: 1.0, y: 3.0, w: 1.4, h: 1.4,
    fontFace: F.header, italic: true, fontSize: 36, bold: true, color: C.accent,
    align: 'center', valign: 'middle', margin: 0,
  });
  s.addText('김지영 (32)', {
    x: 2.6, y: 3.05, w: 2.8, h: 0.45,
    fontFace: F.body, fontSize: 18, bold: true, color: C.fg, margin: 0,
  });
  s.addText('IT 회사 · 강아지 보호자 5년차', {
    x: 2.6, y: 3.5, w: 2.8, h: 0.35,
    fontFace: F.body, fontSize: 11, color: C.muted, margin: 0,
  });
  s.addText('PROFILE', {
    x: 1.0, y: 4.6, w: 4, h: 0.25,
    fontFace: F.body, fontSize: 9, bold: true, color: C.accent, charSpacing: 3, margin: 0,
  });
  s.addText([
    { text: '· 푸들 ‘콩이’ (5살, 5.2kg)', options: { breakLine: true } },
    { text: '· 월 펫 지출 12~18만원', options: { breakLine: true } },
    { text: '· 모바일 70% / 데스크톱 30%' },
  ], {
    x: 1.0, y: 4.9, w: 4, h: 1.6,
    fontFace: F.body, fontSize: 12, color: C.fg, paraSpaceAfter: 6, margin: 0,
  });

  // 우측 — 인용
  drawCard(s, 6.0, 2.7, 6.7, 4.0, { fill: C.sageSoft, border: C.sage });
  s.addText('"', {
    x: 6.2, y: 2.7, w: 1, h: 1.2,
    fontFace: F.header, fontSize: 96, bold: true, color: C.sage, margin: 0,
  });
  s.addText('우리 콩이는 알러지가 있어서 사료를 신중하게 골라야 해요. 사료가 떨어질 때마다 처음부터 다시 검색하는 게 너무 번거로워요. 한 번에 정리된 정보를 보고 싶고, 매달 알아서 배송됐으면 좋겠어요.', {
    x: 6.4, y: 3.7, w: 6.1, h: 2.0,
    fontFace: F.body, italic: true, fontSize: 14, color: C.fg, margin: 0,
  });
  s.addText('— 김지영 님 · 인터뷰 발췌', {
    x: 6.4, y: 5.9, w: 6.1, h: 0.3,
    fontFace: F.body, fontSize: 10, color: C.muted, charSpacing: 2, margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 04 — 핵심 차별화
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'KEY DIFFERENTIATORS');

  s.addText('Pawmart의 5가지 차별화', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('단순 카탈로그가 아닌, 보호자의 일상에 맞춘 큐레이션', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const items = [
    { n: '01', t: 'Pet Friendly 디자인', d: 'Coral · Cream · Sage 펫 도메인 정체성 살린 디자인 시스템 (디자인 토큰 기반)' },
    { n: '02', t: '맞춤 추천', d: '종·연령·체중·건강 상태 → 영양·간식·생활용품을 30초 큐레이션' },
    { n: '03', t: '정기배송', d: '한 번 설정하면 매달 자동 배송 + 상시 10% 할인. 언제든 해지·변경' },
    { n: '04', t: '케어 가이드', d: '입양·영양·환절기 등 전문가 검증 가이드 + 관련 상품 연결' },
    { n: '05', t: '관리자 대시보드', d: '주문 상태 분포, 재고 인라인 수정, 매출 통계까지 한 화면에' },
  ];
  // 5 카드 — 1줄 5분할
  const cardW = 2.4;
  const gap = 0.15;
  const startX = 0.6;
  const startY = 2.8;
  items.forEach((item, i) => {
    const x = startX + i * (cardW + gap);
    drawCard(s, x, startY, cardW, 3.7);
    drawAccentBar(s, x, startY, 3.7, i % 2 === 0 ? C.accent : C.sage);
    s.addText(item.n, {
      x: x + 0.2, y: startY + 0.2, w: 0.8, h: 0.4,
      fontFace: F.header, italic: true, bold: true,
      fontSize: 26, color: i % 2 === 0 ? C.accent : C.sage, margin: 0,
    });
    s.addText(item.t, {
      x: x + 0.2, y: startY + 0.85, w: cardW - 0.4, h: 1.0,
      fontFace: F.body, fontSize: 14, bold: true, color: C.fg, margin: 0,
    });
    s.addText(item.d, {
      x: x + 0.2, y: startY + 1.85, w: cardW - 0.4, h: 1.7,
      fontFace: F.body, fontSize: 10, color: C.muted, paraSpaceAfter: 4, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 05 — User 기능 맵
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'USER FEATURES');

  s.addText('사용자 기능 맵', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('일반 회원이 사용할 수 있는 24개 페이지 · 10개 도메인', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const features = [
    { t: '인증', d: '가입 · 로그인 · 카카오/네이버 SSO 진입점\nJWT 기반 세션 + 자동 갱신' },
    { t: '상품 카탈로그', d: '디바운스 검색 · 다중 필터 · 정렬\n페이징 / 펫·카테고리·뱃지·가격대' },
    { t: '상품 상세', d: '이미지 갤러리 · 효능 배지 · 리뷰\n관련 상품 자동 추천' },
    { t: '장바구니/위시', d: '재고 동시성 안전 · 수량 조정\n위시리스트 · 일괄 처리' },
    { t: '주문 · 결제', d: '주소록 관리 · Toss Payments\n결제 성공/실패/취소 분기' },
    { t: '마이페이지', d: '주문 내역 · 펫 프로필 등록\n프로필 수정 · 회원 탈퇴' },
    { t: '리뷰', d: '포토 후기 (다중 이미지 업로드)\n별점 · 도움된 수' },
    { t: '케어 가이드', d: '에디토리얼 콘텐츠\n태그·펫타입 필터링' },
  ];

  const cols = 4;
  const cardW = 3.0;
  const cardH = 2.0;
  const gapX = 0.13;
  const gapY = 0.2;
  const startX = 0.6;
  const startY = 2.7;
  features.forEach((f, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    drawCard(s, x, y, cardW, cardH);
    s.addShape('ellipse', {
      x: x + 0.25, y: y + 0.3, w: 0.4, h: 0.4,
      fill: { color: C.accentSoft }, line: { color: C.accentSoft, width: 0 },
    });
    s.addText(String(i + 1).padStart(2, '0'), {
      x: x + 0.25, y: y + 0.3, w: 0.4, h: 0.4,
      fontFace: F.body, fontSize: 11, bold: true, color: C.accent,
      align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(f.t, {
      x: x + 0.8, y: y + 0.3, w: cardW - 1, h: 0.4,
      fontFace: F.body, fontSize: 14, bold: true, color: C.fg, margin: 0,
    });
    s.addText(f.d, {
      x: x + 0.25, y: y + 0.85, w: cardW - 0.5, h: 1.0,
      fontFace: F.body, fontSize: 10, color: C.muted, paraSpaceAfter: 2, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 06 — Admin 기능 맵
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'ADMIN FEATURES');

  s.addText('관리자 기능 맵', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('Role-based 가드로 보호된 6개 관리자 페이지', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const adminFeatures = [
    { t: '대시보드', d: 'KPI 4종 (매출·주문·회원·상품)\n주문 상태 분포 바 시각화\n처리 필요 알림 + 빠른 액션' },
    { t: '주문 관리', d: '결제완료→배송완료 상태 전이\n주문 상세 모달 (수령인·상품)\n5단계 필터 + 페이징' },
    { t: '상품 관리', d: '검색 · 페이징 · CRUD\n뱃지 · 카테고리 · 펫타입\n이미지 업로드 통합' },
    { t: '상품 등록', d: '단독 페이지 폼\nImageUploader 통합\n실시간 검증 + 에러 처리' },
    { t: '재고 관리', d: '재고 부족(≤10) · 품절 필터\n인라인 수정 + 되돌리기\n낙관적 동시성 제어' },
    { t: '회원 관리', d: '전체 회원 페이징\n역할 / 상태 표시\n가입일 정렬' },
  ];

  const cols = 3;
  const cardW = 4.05;
  const cardH = 2.0;
  const gapX = 0.13;
  const gapY = 0.2;
  const startX = 0.6;
  const startY = 2.7;
  adminFeatures.forEach((f, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    drawCard(s, x, y, cardW, cardH, { fill: C.darkBg, border: C.darkBg });
    drawAccentBar(s, x, y, cardH, C.accent);
    s.addText(String(i + 1).padStart(2, '0'), {
      x: x + 0.25, y: y + 0.25, w: 0.6, h: 0.35,
      fontFace: F.header, italic: true, fontSize: 18, bold: true, color: C.accent, margin: 0,
    });
    s.addText(f.t, {
      x: x + 0.85, y: y + 0.27, w: cardW - 1, h: 0.35,
      fontFace: F.body, fontSize: 15, bold: true, color: 'FFFFFF', margin: 0,
    });
    s.addText(f.d, {
      x: x + 0.25, y: y + 0.85, w: cardW - 0.5, h: 1.0,
      fontFace: F.body, fontSize: 11, color: C.accentSoft, paraSpaceAfter: 2, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 07 — 기술 스택
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'TECH STACK');

  s.addText('기술 스택', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('최신 안정 버전 위주 · 한국 시장 친화 (Pretendard, Toss)', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const cols = [
    {
      title: 'Frontend',
      sub: '브라우저에서 실행',
      items: [
        ['React', '19'],
        ['TypeScript', '5.6'],
        ['Vite', '8'],
        ['Tailwind CSS', 'v4'],
        ['shadcn/ui', 'latest'],
        ['Zustand', '상태 관리'],
        ['axios', 'HTTP'],
        ['lucide-react', '아이콘'],
        ['Pretendard', '한글 폰트'],
        ['Toss SDK', '결제'],
      ],
      color: C.accent,
    },
    {
      title: 'Backend',
      sub: 'Spring 기반 REST API',
      items: [
        ['Spring Boot', '4'],
        ['Java', '21 (LTS)'],
        ['Spring Security', 'JWT'],
        ['Spring Data JPA', 'Hibernate'],
        ['MySQL', '8'],
        ['Flyway', '마이그레이션 12'],
        ['BCrypt', '암호화'],
        ['Lombok', '보일러플레이트'],
        ['Toss Payments', 'PG 통합'],
        ['MockMvc', '통합 테스트'],
      ],
      color: C.sage,
    },
    {
      title: 'DevOps · Tools',
      sub: '개발 / 빌드 / 협업',
      items: [
        ['Git', '버전관리'],
        ['Gradle', '8 (Kotlin DSL)'],
        ['ESLint', 'flat config'],
        ['Prettier', '포맷'],
        ['npm', '패키지 매니저'],
        ['Karpathy 가이드', '코드 원칙'],
        ['Claude Code', 'AI 페어 프로그래밍'],
        ['GitHub', '원격 저장소'],
        ['Postman/curl', 'API 테스트'],
        ['DBeaver', 'DB 관리'],
      ],
      color: C.gold,
    },
  ];

  const colW = 4.05;
  const gapX = 0.13;
  const startX = 0.6;
  const colY = 2.7;
  const colH = 4.4;

  cols.forEach((col, i) => {
    const x = startX + i * (colW + gapX);
    drawCard(s, x, colY, colW, colH);
    drawAccentBar(s, x, colY, colH, col.color);
    s.addText(col.title, {
      x: x + 0.25, y: colY + 0.2, w: colW - 0.5, h: 0.4,
      fontFace: F.header, italic: true, fontSize: 22, bold: true, color: col.color, margin: 0,
    });
    s.addText(col.sub, {
      x: x + 0.25, y: colY + 0.65, w: colW - 0.5, h: 0.3,
      fontFace: F.body, fontSize: 10, color: C.muted, charSpacing: 2, margin: 0,
    });
    col.items.forEach((item, j) => {
      const y = colY + 1.05 + j * 0.32;
      s.addShape('ellipse', {
        x: x + 0.27, y: y + 0.1, w: 0.08, h: 0.08,
        fill: { color: col.color }, line: { color: col.color, width: 0 },
      });
      s.addText([
        { text: item[0], options: { bold: true, fontSize: 11, color: C.fg } },
        { text: '   ' + item[1], options: { fontSize: 9, color: C.muted } },
      ], {
        x: x + 0.45, y: y, w: colW - 0.6, h: 0.3,
        fontFace: F.body, margin: 0, valign: 'middle',
      });
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 08 — 시스템 아키텍처
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'SYSTEM ARCHITECTURE');

  s.addText('시스템 아키텍처', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('SPA + REST API + RDB · JWT 기반 stateless 인증', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  // 박스 4단 다이어그램 (가로)
  const boxes = [
    { t: '브라우저', d: 'Chrome / Safari / Mobile', color: C.sage, sub: 'Client' },
    { t: 'React 19', d: 'Vite 빌드 · Tailwind v4\nZustand · TanStack', color: C.accent, sub: 'Frontend' },
    { t: 'Spring Boot 4', d: 'REST · JWT · JPA\nSecurity · Validation', color: C.darkBg, sub: 'Backend', white: true },
    { t: 'MySQL 8', d: 'Flyway 마이그레이션\n트랜잭션 · 동시성', color: C.gold, sub: 'Database' },
  ];

  const boxW = 2.7;
  const boxH = 2.4;
  const gap = 0.45;
  const startX = 0.6;
  const startY = 3.2;

  boxes.forEach((b, i) => {
    const x = startX + i * (boxW + gap);
    drawCard(s, x, startY, boxW, boxH, { fill: b.color, border: b.color });
    s.addText(b.sub, {
      x: x + 0.2, y: startY + 0.2, w: boxW - 0.4, h: 0.25,
      fontFace: F.body, fontSize: 9, bold: true, color: b.white ? C.accentSoft : 'FFFFFF',
      charSpacing: 4, margin: 0,
    });
    s.addText(b.t, {
      x: x + 0.2, y: startY + 0.55, w: boxW - 0.4, h: 0.5,
      fontFace: F.header, italic: true, fontSize: 22, bold: true,
      color: 'FFFFFF', margin: 0,
    });
    s.addText(b.d, {
      x: x + 0.2, y: startY + 1.2, w: boxW - 0.4, h: 1.1,
      fontFace: F.body, fontSize: 11, color: 'FFFFFF', margin: 0, paraSpaceAfter: 2,
    });

    // 화살표
    if (i < boxes.length - 1) {
      const arrowX = x + boxW + 0.05;
      s.addShape('rightTriangle', {
        x: arrowX, y: startY + boxH / 2 - 0.15, w: 0.35, h: 0.3,
        fill: { color: C.muted }, line: { color: C.muted, width: 0 },
        rotate: 90,
      });
    }
  });

  // 하단 캡션
  drawCard(s, 0.6, 6.0, 12.1, 1.0, { fill: C.accentSoft, border: C.accent });
  s.addText('통신 흐름', {
    x: 0.85, y: 6.1, w: 3, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.accent, charSpacing: 3, margin: 0,
  });
  s.addText('GET/POST · Bearer JWT · JSON  →  Controller → Service → Repository → JPA → MySQL', {
    x: 0.85, y: 6.45, w: 11.6, h: 0.5,
    fontFace: F.body, fontSize: 12, color: C.fg, margin: 0,
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 09 — 보안·품질
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'SECURITY · QUALITY');

  s.addText('보안 · 품질 · 안정성', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('포트폴리오가 아닌 실서비스 수준의 방어 적용', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const items = [
    { t: 'JWT + Role-based Authz', d: '백엔드 SecurityConfig에 /api/admin/** → hasRole(ADMIN)\n프론트 AdminRoute로 이중 가드' },
    { t: '비밀번호 암호화', d: 'BCrypt(strength 10) · 솔트 자동\n로그인 응답에 password 필드 절대 미포함' },
    { t: 'OWASP Top 10 방어', d: 'XSS (React 자동 escape · DOMPurify)\nSQL Injection (JPA 파라미터 바인딩)\nCSRF (SameSite + JWT 헤더)' },
    { t: '동시성 제어', d: 'Optimistic Lock (@Version)\n재고 차감 시 충돌 감지 → 재시도\nMockMvc 통합 테스트로 시나리오 검증' },
    { t: '예외 처리', d: '@RestControllerAdvice 통합 핸들러\nHTTP 상태 + 에러 코드 일관 응답\n프론트 axios 인터셉터로 자동 토스트' },
    { t: '입력 검증', d: 'Jakarta Validation (@Valid)\nDTO 단계에서 차단\n프론트 zod-style 폼 가드' },
  ];

  const cols = 3;
  const cardW = 4.05;
  const cardH = 2.0;
  const gapX = 0.13;
  const gapY = 0.2;
  const startX = 0.6;
  const startY = 2.7;

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    drawCard(s, x, y, cardW, cardH);
    s.addShape('ellipse', {
      x: x + 0.25, y: y + 0.3, w: 0.5, h: 0.5,
      fill: { color: C.sageSoft }, line: { color: C.sage, width: 1.5 },
    });
    s.addText('✓', {
      x: x + 0.25, y: y + 0.3, w: 0.5, h: 0.5,
      fontFace: F.body, fontSize: 18, bold: true, color: C.sage,
      align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(item.t, {
      x: x + 0.95, y: y + 0.3, w: cardW - 1.1, h: 0.5,
      fontFace: F.body, fontSize: 13, bold: true, color: C.fg, margin: 0,
    });
    s.addText(item.d, {
      x: x + 0.25, y: y + 0.95, w: cardW - 0.5, h: 1.0,
      fontFace: F.body, fontSize: 10, color: C.muted, paraSpaceAfter: 2, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 10 — 디자인 시스템
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'DESIGN SYSTEM');

  s.addText('디자인 시스템 — Pet Friendly Soft', {
    x: 0.6, y: 1.0, w: 11, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('도메인 정체성 + 한국 사용자 친화 + 2026 트렌드 균형', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  // 좌 — 컬러 팔레트
  s.addText('COLOR PALETTE', {
    x: 0.6, y: 2.7, w: 5, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.muted, charSpacing: 4, margin: 0,
  });
  const palette = [
    { c: C.accent, name: 'Soft Coral', hex: '#E07856', use: 'CTA · 강조' },
    { c: C.bg, name: 'Peach Cream', hex: '#FBF7F1', use: '배경' },
    { c: C.sage, name: 'Sage Green', hex: '#88A876', use: '보조 · 친환경' },
    { c: C.fg, name: 'Warm Black', hex: '#1F1916', use: '본문 · 헤더' },
    { c: C.gold, name: 'Honey Gold', hex: '#D4A24A', use: '별점 · 배지' },
  ];
  palette.forEach((p, i) => {
    const y = 3.05 + i * 0.7;
    s.addShape('roundRect', {
      x: 0.6, y, w: 1.0, h: 0.6,
      fill: { color: p.c }, line: { color: C.line, width: 0.5 }, rectRadius: 0.08,
    });
    s.addText(p.name, {
      x: 1.8, y: y + 0.05, w: 2.5, h: 0.3,
      fontFace: F.body, fontSize: 12, bold: true, color: C.fg, margin: 0,
    });
    s.addText(p.hex, {
      x: 1.8, y: y + 0.32, w: 2.5, h: 0.25,
      fontFace: 'Consolas', fontSize: 10, color: C.muted, margin: 0,
    });
    s.addText(p.use, {
      x: 4.4, y: y + 0.18, w: 1.6, h: 0.3,
      fontFace: F.body, fontSize: 10, color: C.muted, italic: true, margin: 0,
    });
  });

  // 우 — 디자인 원칙
  drawCard(s, 6.5, 2.7, 6.2, 4.0);
  s.addText('PRINCIPLES', {
    x: 6.7, y: 2.85, w: 5.8, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.accent, charSpacing: 4, margin: 0,
  });
  const principles = [
    ['모바일 퍼스트', '360 → 768 → 1920 순서로 break point. 한국 모바일 트래픽 70%+ 대응'],
    ['Bento Grid Hero', '큰 메인 + 작은 보조 카드 — 2026 e-commerce 트렌드'],
    ['디자인 토큰', 'CSS custom properties. .dark 모드 자동 대응'],
    ['마이크로 인터랙션', 'hover-lift, 펄스 링, 호버 스케일 — 살아있는 느낌'],
    ['접근성', 'aria-label · WCAG AA 명도 대비 4.5:1+ · prefers-reduced-motion'],
  ];
  principles.forEach((p, i) => {
    const y = 3.25 + i * 0.65;
    s.addShape('ellipse', {
      x: 6.75, y: y + 0.08, w: 0.25, h: 0.25,
      fill: { color: C.accent }, line: { color: C.accent, width: 0 },
    });
    s.addText(String(i + 1), {
      x: 6.75, y: y + 0.08, w: 0.25, h: 0.25,
      fontFace: F.body, fontSize: 10, bold: true, color: 'FFFFFF',
      align: 'center', valign: 'middle', margin: 0,
    });
    s.addText(p[0], {
      x: 7.15, y: y, w: 5.4, h: 0.3,
      fontFace: F.body, fontSize: 13, bold: true, color: C.fg, margin: 0,
    });
    s.addText(p[1], {
      x: 7.15, y: y + 0.27, w: 5.4, h: 0.4,
      fontFace: F.body, fontSize: 10, color: C.muted, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 11 — 개발 프로세스
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'DEV PROCESS');

  s.addText('개발 프로세스', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('Andrej Karpathy 가이드 적용 + 단계별 commit 진화', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  // Karpathy 4 원칙 카드
  const principles = [
    { n: '01', t: 'Think Before Coding', d: '가정 표면화 · 모호함 명시' },
    { n: '02', t: 'Simplicity First', d: '요청 외 기능 X · 추상 X' },
    { n: '03', t: 'Surgical Changes', d: '인접 코드 안 건드림 · 사전 dead code mention만' },
    { n: '04', t: 'Goal-Driven', d: '검증 가능한 성공 기준 정의' },
  ];

  principles.forEach((p, i) => {
    const x = 0.6 + i * 3.05;
    const y = 2.7;
    drawCard(s, x, y, 2.95, 1.7);
    s.addText(p.n, {
      x: x + 0.2, y: y + 0.2, w: 0.6, h: 0.4,
      fontFace: F.header, italic: true, fontSize: 22, bold: true, color: C.accent, margin: 0,
    });
    s.addText(p.t, {
      x: x + 0.85, y: y + 0.25, w: 2.0, h: 0.4,
      fontFace: F.body, fontSize: 13, bold: true, color: C.fg, margin: 0,
    });
    s.addText(p.d, {
      x: x + 0.2, y: y + 0.85, w: 2.55, h: 0.7,
      fontFace: F.body, fontSize: 10, color: C.muted, margin: 0,
    });
  });

  // 커밋 타임라인
  s.addText('COMMIT HISTORY', {
    x: 0.6, y: 4.7, w: 5, h: 0.3,
    fontFace: F.body, fontSize: 10, bold: true, color: C.muted, charSpacing: 4, margin: 0,
  });

  // 저장소 비공개 → 확인 불가능한 커밋 해시 대신 날짜로 표시, 실제 커밋 순서(시간순)대로 정렬
  const commits = [
    { h: '04.24', d: 'Spring Boot 4 + React 19 기반 리뉴얼 초기 구성' },
    { h: '04.24', d: '보안·예외처리·동시성 방어 강화 + MockMvc 통합테스트' },
    { h: '04.26', d: 'Pawmart 전체 기능 구현 완료' },
    { h: '04.27', d: '재고 동시성 제어 + OrderService 통합테스트' },
    { h: '05.04', d: 'Pet Friendly 디자인 + admin 권한·대시보드 개편' },
  ];

  // 가로 타임라인 라인
  s.addShape('rect', {
    x: 0.7, y: 5.85, w: 12.0, h: 0.02,
    fill: { color: C.line }, line: { color: C.line, width: 0 },
  });

  commits.forEach((c, i) => {
    const x = 0.6 + i * 2.45;
    s.addShape('ellipse', {
      x: x + 0.65, y: 5.7, w: 0.3, h: 0.3,
      fill: { color: i === commits.length - 1 ? C.accent : C.card },
      line: { color: C.accent, width: 2 },
    });
    s.addText(c.h, {
      x: x, y: 5.2, w: 1.6, h: 0.3,
      fontFace: 'Consolas', fontSize: 10, bold: true, color: C.fg, align: 'center', margin: 0,
    });
    s.addText(c.d, {
      x: x - 0.3, y: 6.15, w: 2.2, h: 0.8,
      fontFace: F.body, fontSize: 9, color: C.muted, align: 'center', margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 12 — 프로젝트 스케일 (통계)
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'PROJECT SCALE');

  s.addText('프로젝트 규모', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('숫자로 보는 Pawmart', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const stats = [
    { n: '24+', l: '페이지', d: '사용자 18 + 관리자 6' },
    { n: '60+', l: '컴포넌트', d: '재사용 가능한 UI 단위' },
    { n: '12', l: 'API 모듈', d: 'frontend/src/api/*.ts' },
    { n: '12', l: 'DB 마이그레이션', d: 'Flyway 자동 적용' },
    { n: '3,176', l: '추가 라인', d: '최근 디자인 개편 (+5%)' },
    { n: '10+', l: '엔티티', d: 'Member·Product·Order·Cart·Pet…' },
  ];

  // 3x2 큰 통계 카드
  const cardW = 4.05;
  const cardH = 2.0;
  const gapX = 0.13;
  const gapY = 0.2;
  const startX = 0.6;
  const startY = 2.7;

  stats.forEach((stat, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = startX + col * (cardW + gapX);
    const y = startY + row * (cardH + gapY);
    const isAccent = i === 0 || i === 4;
    drawCard(s, x, y, cardW, cardH, {
      fill: isAccent ? C.accent : C.card,
      border: isAccent ? C.accent : C.line,
    });
    s.addText(stat.n, {
      x: x + 0.25, y: y + 0.25, w: cardW - 0.5, h: 0.95,
      fontFace: F.header, italic: true, fontSize: 60, bold: true,
      color: isAccent ? 'FFFFFF' : C.fg, margin: 0,
    });
    s.addText(stat.l, {
      x: x + 0.25, y: y + 1.25, w: cardW - 0.5, h: 0.35,
      fontFace: F.body, fontSize: 13, bold: true,
      color: isAccent ? 'FFFFFF' : C.fg, charSpacing: 2, margin: 0,
    });
    s.addText(stat.d, {
      x: x + 0.25, y: y + 1.6, w: cardW - 0.5, h: 0.3,
      fontFace: F.body, fontSize: 10, color: isAccent ? C.accentSoft : C.muted, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 13 — 디자인 R&D
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'DESIGN R&D');

  s.addText('디자인 시안 — 4종 비교 후 결정', {
    x: 0.6, y: 1.0, w: 11, h: 0.7,
    fontFace: F.header, fontSize: 36, bold: true, color: C.fg, margin: 0,
  });
  s.addText('한 가지 안을 단번에 만들지 않고, 4개 컨셉을 만들어 검증', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const options = [
    { n: 'A', t: 'Minimal Editorial', d: '흑백 + 1 accent\n무신사 / 29CM 톤', color: 'FFFFFF', fg: '0A0A0A', selected: false },
    { n: 'B', t: 'Warm Marketplace', d: '크림+brown+sage\n마켓컬리 / Chewy', color: 'FAF6EE', fg: '5B4023', selected: false },
    { n: 'C', t: 'Playful Pet', d: '파스텔, 둥근 카드\nPetsmart 풍', color: 'FFE3D6', fg: '3D2C2A', selected: false },
    { n: 'D', t: 'Modern Korean Pet', d: 'B+C 하이브리드\n선정 — 포트폴리오 + 한국 사용자', color: C.accent, fg: 'FFFFFF', selected: true },
  ];

  const cardW = 3.0;
  const cardH = 4.2;
  const gapX = 0.13;
  const startX = 0.6;
  const startY = 2.7;

  options.forEach((opt, i) => {
    const x = startX + i * (cardW + gapX);
    drawCard(s, x, startY, cardW, cardH, {
      fill: opt.color,
      border: opt.selected ? C.accent : C.line,
    });

    if (opt.selected) {
      // 선정 라벨 우상단
      s.addShape('roundRect', {
        x: x + cardW - 1.05, y: startY + 0.2, w: 0.9, h: 0.35,
        fill: { color: C.darkBg }, line: { color: C.darkBg, width: 0 }, rectRadius: 0.05,
      });
      s.addText('SELECTED', {
        x: x + cardW - 1.05, y: startY + 0.2, w: 0.9, h: 0.35,
        fontFace: F.body, fontSize: 9, bold: true, color: 'FFFFFF',
        align: 'center', valign: 'middle', charSpacing: 2, margin: 0,
      });
    }

    s.addText(opt.n, {
      x: x + 0.3, y: startY + 0.3, w: 0.8, h: 0.6,
      fontFace: F.header, italic: true, fontSize: 48, bold: true, color: opt.fg, margin: 0,
    });

    // 미니어처 색상 블록 (시각화)
    s.addShape('rect', {
      x: x + 0.3, y: startY + 1.4, w: cardW - 0.6, h: 1.5,
      fill: { color: opt.color === 'FFFFFF' ? '1A1A1A' : opt.color },
      line: { color: C.line, width: 1 },
    });
    if (opt.selected) {
      // 시안 D는 미니어처 안에 컬러 strip
      s.addShape('rect', {
        x: x + 0.3, y: startY + 2.0, w: cardW - 0.6, h: 0.3,
        fill: { color: 'FBF7F1' }, line: { color: 'FBF7F1', width: 0 },
      });
      s.addShape('rect', {
        x: x + 0.3, y: startY + 2.3, w: 0.5, h: 0.3,
        fill: { color: '88A876' }, line: { color: '88A876', width: 0 },
      });
      s.addShape('rect', {
        x: x + 0.8, y: startY + 2.3, w: 0.5, h: 0.3,
        fill: { color: 'E07856' }, line: { color: 'E07856', width: 0 },
      });
    }

    s.addText(opt.t, {
      x: x + 0.3, y: startY + 3.05, w: cardW - 0.6, h: 0.4,
      fontFace: F.body, fontSize: 14, bold: true, color: opt.fg, margin: 0,
    });
    s.addText(opt.d, {
      x: x + 0.3, y: startY + 3.5, w: cardW - 0.6, h: 0.6,
      fontFace: F.body, fontSize: 10,
      color: opt.fg === 'FFFFFF' ? C.accentSoft : C.muted,
      margin: 0, paraSpaceAfter: 2,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 14 — 하이라이트 화면
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'KEY SCREENS');

  s.addText('하이라이트 화면 3선', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('홈 / 상품 상세 / 관리자 대시보드 — 디자인 일관성 확인', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const screens = [
    {
      t: 'HOMEPAGE',
      d: 'Bento Hero + 카테고리 + 맞춤추천\n+ 베스트 + 정기배송 + 케어가이드',
      color: C.accent,
    },
    {
      t: 'PRODUCT DETAIL',
      d: '이미지 갤러리 + 효능 배지\n+ 리뷰 + 관련상품',
      color: C.sage,
    },
    {
      t: 'ADMIN DASHBOARD',
      d: 'KPI 4종 + 주문 상태 분포 바\n+ 최근주문 + 빠른액션',
      color: C.darkBg,
    },
  ];

  const cardW = 4.05;
  const cardH = 4.4;
  const gapX = 0.13;
  const startX = 0.6;
  const startY = 2.7;

  screens.forEach((scr, i) => {
    const x = startX + i * (cardW + gapX);
    // 외곽 카드
    drawCard(s, x, startY, cardW, cardH);
    // 상단 컬러 헤더
    s.addShape('rect', {
      x: x, y: startY, w: cardW, h: 0.8,
      fill: { color: scr.color }, line: { color: scr.color, width: 0 },
    });
    s.addText(scr.t, {
      x: x + 0.25, y: startY + 0.2, w: cardW - 0.5, h: 0.4,
      fontFace: F.body, fontSize: 13, bold: true, color: 'FFFFFF',
      charSpacing: 4, margin: 0,
    });

    // 미니어처 영역 (모형)
    const mockX = x + 0.25;
    const mockY = startY + 1.0;
    const mockW = cardW - 0.5;
    const mockH = 2.5;
    s.addShape('roundRect', {
      x: mockX, y: mockY, w: mockW, h: mockH,
      fill: { color: C.bg }, line: { color: C.line, width: 1 }, rectRadius: 0.05,
    });
    // 모형 내부 (간단 박스들)
    s.addShape('rect', {
      x: mockX + 0.15, y: mockY + 0.15, w: mockW - 0.3, h: 0.18,
      fill: { color: scr.color }, line: { color: scr.color, width: 0 },
    });
    if (i === 0) {
      // 홈: 큰 hero + 작은 두 카드
      s.addShape('rect', {
        x: mockX + 0.15, y: mockY + 0.45, w: (mockW - 0.4) * 0.6, h: 1.0,
        fill: { color: C.accentSoft }, line: { color: C.accent, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + (mockW - 0.4) * 0.6 + 0.25, y: mockY + 0.45, w: (mockW - 0.4) * 0.4 - 0.05, h: 0.45,
        fill: { color: C.sage }, line: { color: C.sage, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + (mockW - 0.4) * 0.6 + 0.25, y: mockY + 1.0, w: (mockW - 0.4) * 0.4 - 0.05, h: 0.45,
        fill: { color: C.accent }, line: { color: C.accent, width: 0 },
      });
      // 그리드 카드 4개
      for (let k = 0; k < 4; k++) {
        s.addShape('rect', {
          x: mockX + 0.15 + k * (mockW - 0.4) / 4, y: mockY + 1.6, w: (mockW - 0.4) / 4 - 0.05, h: 0.7,
          fill: { color: C.card }, line: { color: C.line, width: 1 },
        });
      }
    } else if (i === 1) {
      // 상품상세: 좌 이미지 + 우 정보
      s.addShape('rect', {
        x: mockX + 0.15, y: mockY + 0.45, w: (mockW - 0.4) * 0.5, h: 1.5,
        fill: { color: C.sageSoft }, line: { color: C.sage, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + 0.15 + (mockW - 0.4) * 0.5 + 0.1, y: mockY + 0.45, w: (mockW - 0.4) * 0.5 - 0.1, h: 0.25,
        fill: { color: C.fg }, line: { color: C.fg, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + 0.15 + (mockW - 0.4) * 0.5 + 0.1, y: mockY + 0.8, w: (mockW - 0.4) * 0.5 - 0.1, h: 0.18,
        fill: { color: C.muted }, line: { color: C.muted, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + 0.15 + (mockW - 0.4) * 0.5 + 0.1, y: mockY + 1.6, w: (mockW - 0.4) * 0.5 - 0.1, h: 0.35,
        fill: { color: C.accent }, line: { color: C.accent, width: 0 },
      });
      // 리뷰 thumbnails
      for (let k = 0; k < 3; k++) {
        s.addShape('rect', {
          x: mockX + 0.15 + k * (mockW - 0.4) / 3, y: mockY + 2.05, w: (mockW - 0.4) / 3 - 0.05, h: 0.3,
          fill: { color: C.card }, line: { color: C.line, width: 1 },
        });
      }
    } else {
      // Admin: 4 KPI + 차트 + 테이블
      for (let k = 0; k < 4; k++) {
        s.addShape('rect', {
          x: mockX + 0.15 + k * ((mockW - 0.4) / 4), y: mockY + 0.45, w: (mockW - 0.4) / 4 - 0.05, h: 0.6,
          fill: { color: C.card }, line: { color: C.line, width: 1 },
        });
      }
      // 분포 바
      s.addShape('rect', {
        x: mockX + 0.15, y: mockY + 1.2, w: (mockW - 0.4) * 0.55, h: 0.18,
        fill: { color: C.accent }, line: { color: C.accent, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + 0.15 + (mockW - 0.4) * 0.55, y: mockY + 1.2, w: (mockW - 0.4) * 0.25, h: 0.18,
        fill: { color: C.sage }, line: { color: C.sage, width: 0 },
      });
      s.addShape('rect', {
        x: mockX + 0.15 + (mockW - 0.4) * 0.8, y: mockY + 1.2, w: (mockW - 0.4) * 0.2, h: 0.18,
        fill: { color: C.gold }, line: { color: C.gold, width: 0 },
      });
      // 테이블 행
      for (let k = 0; k < 4; k++) {
        s.addShape('rect', {
          x: mockX + 0.15, y: mockY + 1.55 + k * 0.2, w: mockW - 0.4, h: 0.15,
          fill: { color: k % 2 === 0 ? C.card : C.bg }, line: { color: C.line, width: 0.5 },
        });
      }
    }

    s.addText(scr.d, {
      x: x + 0.25, y: startY + 3.55, w: cardW - 0.5, h: 0.85,
      fontFace: F.body, fontSize: 11, color: C.muted, margin: 0,
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 15 — 로드맵
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.bg };
  drawShell(s, 'ROADMAP');

  s.addText('향후 로드맵', {
    x: 0.6, y: 1.0, w: 10, h: 0.7,
    fontFace: F.header, fontSize: 38, bold: true, color: C.fg, margin: 0,
  });
  s.addText('지금은 v1. 다음 단계도 명확하게 정의돼 있다.', {
    x: 0.6, y: 1.85, w: 11, h: 0.4,
    fontFace: F.body, fontSize: 15, color: C.muted, margin: 0,
  });

  const phases = [
    {
      tag: 'NEAR · Q2',
      title: '단기',
      color: C.accent,
      items: [
        '대시보드 매출 추이 차트 (recharts)',
        'E2E 테스트 (Playwright)',
        '잔여 이모지 정리 + 디자인 토큰 일관성',
        '관리자 권한 거부 토스트 안내',
      ],
    },
    {
      tag: 'MID · Q3',
      title: '중기',
      color: C.sage,
      items: [
        'Docker Compose 자동화 (백엔드+DB+Adminer)',
        '결제 다양화 (네이버페이 · 카카오페이)',
        '품절 임박 자동 알림 (관리자)',
        '리뷰 자동 모더레이션 (욕설 필터)',
      ],
    },
    {
      tag: 'LONG · Q4+',
      title: '장기',
      color: C.gold,
      items: [
        'AI 기반 상품 추천 고도화 (행동 데이터 기반)',
        '모바일 앱 (React Native 또는 Flutter)',
        '정기배송 데이터 분석 대시보드',
        'B2B 채널 (펫샵 · 동물병원)',
      ],
    },
  ];

  const cardW = 4.05;
  const cardH = 4.4;
  const gapX = 0.13;
  const startX = 0.6;
  const startY = 2.7;

  phases.forEach((p, i) => {
    const x = startX + i * (cardW + gapX);
    drawCard(s, x, startY, cardW, cardH);
    drawAccentBar(s, x, startY, cardH, p.color);
    s.addShape('roundRect', {
      x: x + 0.25, y: startY + 0.25, w: 1.7, h: 0.4,
      fill: { color: p.color }, line: { color: p.color, width: 0 }, rectRadius: 0.05,
    });
    s.addText(p.tag, {
      x: x + 0.25, y: startY + 0.25, w: 1.7, h: 0.4,
      fontFace: F.body, fontSize: 9, bold: true, color: 'FFFFFF',
      align: 'center', valign: 'middle', charSpacing: 3, margin: 0,
    });
    s.addText(p.title, {
      x: x + 0.25, y: startY + 0.85, w: cardW - 0.5, h: 0.5,
      fontFace: F.header, italic: true, fontSize: 26, bold: true, color: C.fg, margin: 0,
    });
    p.items.forEach((it, j) => {
      const y = startY + 1.6 + j * 0.6;
      s.addShape('ellipse', {
        x: x + 0.3, y: y + 0.1, w: 0.15, h: 0.15,
        fill: { color: p.color }, line: { color: p.color, width: 0 },
      });
      s.addText(it, {
        x: x + 0.55, y: y, w: cardW - 0.7, h: 0.5,
        fontFace: F.body, fontSize: 11, color: C.fg, margin: 0,
      });
    });
  });
}

// ════════════════════════════════════════════════════════════
// SLIDE 16 — Thank you
// ════════════════════════════════════════════════════════════
{
  const s = pres.addSlide();
  s.background = { color: C.darkBg };

  // 좌상단 라벨
  s.addShape('ellipse', { x: 0.6, y: 0.42, w: 0.12, h: 0.12, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });
  s.addText('END · THANK YOU', {
    x: 0.85, y: 0.36, w: 5, h: 0.25,
    fontFace: F.body, fontSize: 9, color: 'FFFFFF', bold: true, charSpacing: 5, margin: 0,
  });
  s.addText('16 / 16', {
    x: 11.8, y: 0.36, w: 1.2, h: 0.25,
    fontFace: F.body, fontSize: 9, color: 'FFFFFF', bold: true, align: 'right', charSpacing: 3, margin: 0,
  });

  // 컬러 블록 좌측 (vertical)
  s.addShape('rect', { x: 0, y: 0, w: 0.5, h: 7.5, fill: { color: C.accent }, line: { color: C.accent, width: 0 } });

  // 메인 메시지
  s.addText('Thank You.', {
    x: 1.0, y: 2.3, w: 11, h: 1.2,
    fontFace: F.header, italic: true, fontSize: 88, bold: true, color: 'FFFFFF', margin: 0,
  });
  s.addText('함께 만들어 갈 수 있는 기회가 되길 바라며.', {
    x: 1.0, y: 3.6, w: 11, h: 0.5,
    fontFace: F.body, fontSize: 18, color: C.accentSoft, margin: 0,
  });

  // 연락처 카드 (가로)
  drawCard(s, 1.0, 5.0, 11.5, 1.6, { fill: 'FFFFFF', border: 'FFFFFF' });
  const contacts = [
    { l: '이름', v: '서인석' },
    // TODO: 실제 연락받을 이메일로 교체
    { l: '이메일', v: 'seok2@pawmart.kr' },
    // 프로젝트 저장소는 비공개 → 공개 포트폴리오 사이트로 연결
    { l: '포트폴리오', v: 'budurang22.github.io' },
    { l: '포지션', v: 'Backend · DevOps' },
  ];
  contacts.forEach((c, i) => {
    const x = 1.0 + i * 2.875;
    s.addText(c.l, {
      x: x + 0.3, y: 5.2, w: 2.5, h: 0.3,
      fontFace: F.body, fontSize: 9, bold: true, color: C.accent, charSpacing: 3, margin: 0,
    });
    s.addText(c.v, {
      x: x + 0.3, y: 5.55, w: 2.5, h: 0.9,
      fontFace: F.body, fontSize: 13, bold: true, color: C.fg, margin: 0,
    });
  });

  // 푸터
  s.addText('PAWMART · PORTFOLIO 2026', {
    x: 1.0, y: 7.0, w: 6, h: 0.3,
    fontFace: F.body, fontSize: 9, color: C.accentSoft, charSpacing: 5, margin: 0,
  });
  s.addText('Crafted with React 19, Spring Boot 4 — and care.', {
    x: 7.0, y: 7.0, w: 5.5, h: 0.3,
    fontFace: F.body, italic: true, fontSize: 9, color: C.accentSoft, align: 'right', margin: 0,
  });
}

// ── 저장 ──────────────────────────────────────────────────────
pres.writeFile({ fileName: 'pawmart.pptx' }).then((file) => {
  console.log('✅ Generated:', file);
});
