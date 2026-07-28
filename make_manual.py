# -*- coding: utf-8 -*-
"""
직원용 사용설명서 — 단일 파일 인터랙티브 HTML 생성기 (영속 소스 + 재생성)
────────────────────────────────────────────────────────────
· 설명서를 .html 파일이 아니라 이 .py 를 '원본'으로 두고 실행 중에 만들어 낸다.
  이유: (1) 인터넷망 보안프로그램이 특정 형식 파일을 삭제해도 exe 안에 들어 있어 안전
        (2) 설명서가 지워지거나 실수로 삭제돼도 실행하면 항상 되살아남
        (3) 업무망(파이썬 없음)에서도 exe 하나로 동작
· 프로그램(트레이/메인창/‘--manual’/최초 실행)이 이 모듈로 HTML을 새로 만들어 연다.
· 실제 스크린샷은 manual_assets.IMAGES(base64)에서 인라인 → 외부 파일 없는 단일 HTML.
      python make_manual.py --open     # 직접 생성+열기(개발용)
"""
import os
import sys
import webbrowser

try:
    from manual_assets import IMAGES          # {"main": data-uri, "rules": data-uri}
except Exception:
    IMAGES = {}
try:
    from manual_assets_static import STATIC    # {"drm": data-uri} 사내 캡처 이미지
except Exception:
    STATIC = {}
try:
    from manual_assets_lottie import PLAYER_B64, ANIM   # Lottie 재생기 + 애니 JSON
except Exception:
    PLAYER_B64, ANIM = "", {}
try:
    from manual_assets_shots import SHOTS    # 실제 화면 캡처(트레이/메뉴/폴더 등)
except Exception:
    SHOTS = {}


def _data_dir():
    """설치형(frozen)은 사용자 '문서' 폴더, 개발 시엔 스크립트 폴더 (앱들과 동일 규칙)."""
    if getattr(sys, "frozen", False):
        return os.path.join(os.path.expanduser("~"), "Documents", "엑셀파일_개인정보_마스킹")
    return os.path.dirname(os.path.abspath(__file__))


HTML = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>엑셀파일 개인정보 마스킹 · 사용설명서</title>
<style>
  :root{
    --navy:#16324f; --blue:#2563eb; --sky:#e8f0fe; --green:#15a34a;
    --red:#dc2626; --amber:#f59e0b; --ink:#1f2937; --muted:#64748b;
    --line:#e6eaf0; --bg:#f5f7fa; --card:#ffffff; --radius:16px;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  html{scroll-behavior:smooth;scroll-padding-top:64px}
  body{font-family:"Pretendard","맑은 고딕","Malgun Gothic",system-ui,sans-serif;
       color:var(--ink);background:var(--bg);line-height:1.7;font-size:16px;
       -webkit-font-smoothing:antialiased}
  .wrap{max-width:900px;margin:0 auto;padding:0 18px 90px}
  a{color:var(--blue)}
  /* HERO */
  header.hero{background:radial-gradient(120% 140% at 0% 0%,#2b5a86 0%,#16324f 60%);
       color:#fff;padding:44px 18px 34px;text-align:center}
  header.hero .lock{font-size:2.4rem}
  header.hero h1{font-size:1.85rem;letter-spacing:-.6px;margin-top:6px}
  header.hero p{opacity:.9;margin-top:10px;font-size:1.02rem}
  /* QUICK START */
  .quick{max-width:900px;margin:-24px auto 0;background:#fff;border-radius:var(--radius);
       padding:20px;box-shadow:0 12px 34px rgba(22,50,79,.14);position:relative}
  .quick h3{color:var(--navy);font-size:1.05rem;margin-bottom:12px}
  .qsteps{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}
  .qstep{background:var(--sky);border-radius:12px;padding:14px;text-align:center}
  .qstep .n{width:26px;height:26px;border-radius:50%;background:var(--blue);color:#fff;
       font-weight:800;display:inline-flex;align-items:center;justify-content:center;
       font-size:.85rem;margin-bottom:8px}
  .qstep b{display:block;color:var(--navy);font-size:.96rem}
  .qstep span{font-size:.84rem;color:var(--muted)}
  /* NAV */
  nav.toc{position:sticky;top:0;z-index:30;background:rgba(255,255,255,.92);
       backdrop-filter:blur(8px);border-bottom:1px solid var(--line);
       padding:9px 0;margin-top:26px;overflow-x:auto;white-space:nowrap;
       -webkit-overflow-scrolling:touch}
  nav.toc .inner{max-width:900px;margin:0 auto;padding:0 14px;display:flex;gap:6px}
  nav.toc a{flex:0 0 auto;font-size:.86rem;color:var(--navy);text-decoration:none;
       padding:7px 13px;border-radius:999px;background:#eef2f7;font-weight:700}
  nav.toc a:hover{background:#dbe4ef}
  /* SECTION */
  section{background:var(--card);border-radius:var(--radius);padding:24px 22px;margin-top:22px;
       box-shadow:0 3px 16px rgba(22,50,79,.05)}
  h2{font-size:1.32rem;color:var(--navy);display:flex;align-items:center;gap:10px}
  h2 .num{background:var(--navy);color:#fff;border-radius:9px;font-size:.85rem;
       padding:3px 10px;flex:0 0 auto}
  .sub{color:var(--muted);font-size:.96rem;margin:6px 0 14px}
  h3{font-size:1.06rem;margin:18px 0 6px;color:#12283f}
  p{margin:8px 0}
  ul.plain{margin:8px 0 8px 20px}
  .grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:12px}
  .mini{background:#f8fafc;border:1px solid var(--line);border-radius:12px;padding:16px}
  .mini .ic{font-size:1.8rem}
  .mini b{display:block;margin:8px 0 3px;color:var(--navy)}
  .mini span{font-size:.9rem;color:var(--muted)}
  table{width:100%;border-collapse:collapse;margin:12px 0;font-size:.96rem}
  th,td{border:1px solid var(--line);padding:10px 12px;text-align:left;vertical-align:top}
  th{background:#f1f5f9;color:var(--navy);white-space:nowrap}
  code.mask{background:#eef2f7;border-radius:6px;padding:1px 8px;font-weight:700;
       letter-spacing:1px;color:#0f172a}
  /* STEP */
  .steps{margin-top:10px}
  .step{position:relative;padding:16px 16px 16px 60px;border:1px solid var(--line);
       border-radius:14px;margin:12px 0;background:#fff}
  .step .k{position:absolute;left:14px;top:14px;width:32px;height:32px;border-radius:50%;
       background:var(--blue);color:#fff;font-weight:800;display:flex;align-items:center;
       justify-content:center}
  .step b{color:var(--navy)}
  /* SCREENSHOT */
  figure.shot{margin:14px 0;border-radius:14px;overflow:hidden;border:1px solid var(--line);
       box-shadow:0 10px 26px rgba(2,6,23,.10)}
  figure.shot img{display:block;width:100%}
  figure.shot figcaption{background:#0f213500;color:var(--muted);font-size:.86rem;
       padding:9px 12px;background:#f8fafc;border-top:1px solid var(--line)}
  .tag-real{display:inline-block;font-size:.72rem;font-weight:800;color:#166534;
       background:#dcfce7;border-radius:999px;padding:1px 8px;margin-right:6px}
  .tag-ill{display:inline-block;font-size:.72rem;font-weight:800;color:#5b21b6;
       background:#ede9fe;border-radius:999px;padding:1px 8px;margin-right:6px}
  /* WINDOW MOCKUP */
  .win{border:1px solid #cbd5e1;border-radius:12px;overflow:hidden;margin:12px 0;
       box-shadow:0 8px 20px rgba(2,6,23,.10);background:#fff}
  .win .bar{background:#e8edf3;padding:8px 12px;font-size:.82rem;color:#475569;
       display:flex;align-items:center;gap:7px}
  .win .dot{width:10px;height:10px;border-radius:50%;background:#cbd5e1}
  .menu{list-style:none}
  .menu li{padding:11px 14px;border-bottom:1px solid #f1f5f9;font-size:.95rem;
       display:flex;justify-content:space-between;align-items:center}
  .menu li:last-child{border-bottom:0}
  .menu li.hi{background:#eff6ff;font-weight:800;color:var(--blue)}
  .menu li .arrow{color:var(--blue);font-size:.82rem}
  .files .row{display:flex;align-items:center;gap:10px;padding:9px 4px;font-size:.98rem}
  .badge{font-size:.78rem;font-weight:800;border-radius:999px;padding:2px 10px;flex:0 0 auto}
  .badge.no{background:#fee2e2;color:var(--red)} .badge.ok{background:#dcfce7;color:#166534}
  /* FOLDER TREE */
  .tree{background:#0f2136;color:#dbe7f5;border-radius:12px;padding:16px 18px;
       font-family:"D2Coding","Consolas",monospace;font-size:.95rem;line-height:2;overflow-x:auto}
  .tree .f{color:#7dd3fc} .tree .c{color:#93c5fd} .tree .note{color:#94a3b8}
  .tree .hot{color:#86efac;font-weight:700}
  /* CALLOUT */
  .callout{border-radius:12px;padding:14px 16px;margin:14px 0;font-size:.97rem}
  .callout.warn{background:#fffbeb;border:1px solid #fde68a;color:#92400e}
  .callout.info{background:#eff6ff;border:1px solid #bfdbfe;color:#1e40af}
  .callout.ok{background:#f0fdf4;border:1px solid #bbf7d0;color:#166534}
  .big{font-size:1.05rem}
  /* CHECKLIST */
  .check li{list-style:none;display:flex;gap:12px;align-items:flex-start;
       padding:12px;border:1px solid var(--line);border-radius:12px;margin:8px 0;
       cursor:pointer;background:#fff;transition:.15s}
  .check li.done{background:#f0fdf4;border-color:#bbf7d0}
  .check input{width:21px;height:21px;margin-top:2px;flex:0 0 auto;cursor:pointer;accent-color:var(--green)}
  .progress{height:10px;background:#e2e8f0;border-radius:999px;overflow:hidden;margin:8px 0 2px}
  .progress > i{display:block;height:100%;width:0;background:var(--green);transition:width .35s}
  /* FAQ */
  .faq details{border:1px solid var(--line);border-radius:12px;margin:8px 0;background:#fff}
  .faq summary{cursor:pointer;padding:13px 15px;font-weight:700;color:var(--navy);list-style:none}
  .faq summary::-webkit-details-marker{display:none}
  .faq summary::before{content:"＋ ";color:var(--blue)}
  .faq details[open] summary::before{content:"－ "}
  .faq .ans{padding:0 15px 15px;color:#334155}
  /* 실제 캡처 주석(번호 핫스팟 + 설명카드) */
  .annot{margin:14px 0}
  .shotwrap{position:relative;display:block;border-radius:12px;overflow:hidden;
       border:1px solid var(--line);box-shadow:0 8px 22px rgba(2,6,23,.10)}
  .shotwrap img{width:100%;display:block}
  .hot{position:absolute;transform:translate(-50%,-50%);width:26px;height:26px;
       border-radius:50%;background:var(--blue);color:#fff;font-weight:800;font-size:.8rem;
       display:flex;align-items:center;justify-content:center;border:2px solid #fff;
       box-shadow:0 0 0 0 rgba(37,99,235,.55);animation:hotpulse 2.2s infinite;z-index:2}
  @keyframes hotpulse{0%{box-shadow:0 0 0 0 rgba(37,99,235,.5)}
       70%{box-shadow:0 0 0 13px rgba(37,99,235,0)}100%{box-shadow:0 0 0 0 rgba(37,99,235,0)}}
  .legend{list-style:none;margin-top:12px;display:grid;gap:7px}
  .legend li{display:flex;gap:11px;align-items:flex-start;background:#f8fafc;
       border:1px solid var(--line);border-radius:10px;padding:10px 12px;font-size:.95rem}
  .legend .ln{flex:0 0 auto;width:24px;height:24px;border-radius:50%;background:var(--navy);
       color:#fff;font-weight:800;font-size:.82rem;display:flex;align-items:center;
       justify-content:center;margin-top:1px}
  .legend li.key{background:#eff6ff;border-color:#bfdbfe}
  .legend li.key .ln{background:var(--blue)}
  .lottie{width:60px;height:60px;display:inline-block;vertical-align:middle;flex:0 0 auto}
  .lottie.sm{width:40px;height:40px} .lottie.lg{width:104px;height:104px}
  .clickrow{display:flex;align-items:center;gap:6px}
  footer{text-align:center;color:var(--muted);font-size:.86rem;margin-top:28px;padding:0 12px}
  @media (max-width:640px){
    body{font-size:15px}
    header.hero h1{font-size:1.4rem}
    .qsteps,.grid3{grid-template-columns:1fr}
    .step{padding-left:54px}
    h2{font-size:1.16rem}
  }
  @media print{ nav.toc,.quick{display:none} section{box-shadow:none;break-inside:avoid} }
</style>
</head>
<body>
<header class="hero">
  <div class="lock">🔒</div>
  <h1>엑셀파일 개인정보 마스킹</h1>
  <p>컴퓨터를 잘 몰라도 누구나 따라 할 수 있는 사용설명서</p>
</header>

<div class="wrap">
  <!-- QUICK START -->
  <div class="quick">
    <h3>⚡ 3초 요약 — 이것만 하면 됩니다</h3>
    <div class="qsteps">
      <div class="qstep"><span class="n">1</span><b>파일 오른쪽 클릭</b><span>엑셀 파일 위에서</span></div>
      <div class="qstep"><span class="n">2</span><b>‘마스킹 사본 만들기’</b><span>메뉴에서 클릭</span></div>
      <div class="qstep"><span class="n">3</span><b>‘_마스킹’ 파일을 FTC로</b><span>새로 생긴 안전한 사본</span></div>
    </div>
  </div>

  <nav class="toc"><div class="inner">
    <a href="#why">왜 하나요?</a>
    <a href="#setup">처음 준비</a>
    <a href="#drm" style="background:#fee2e2;color:#b91c1c">⚠ DRM 해제</a>
    <a href="#use">3단계 사용법</a>
    <a href="#window">도구 창 살펴보기</a>
    <a href="#folder">폴더로 자동 처리</a>
    <a href="#check">보내기 전 확인</a>
    <a href="#rule">항상 가릴 칸</a>
    <a href="#faq">자주 묻는 질문</a>
    <a href="#words">용어 풀이</a>
  </div></nav>

  <!-- WHY -->
  <section id="why">
    <h2><span class="num">1</span> 왜 이걸 해야 하나요?</h2>
    <div class="sub">개인정보가 새어 나가면 고객과 회사 모두가 위험해집니다.</div>
    <div class="grid3">
      <div class="mini"><div class="ic">🔑</div><b>개인정보 = 집 열쇠</b>
        <span>이름·주민번호·전화번호가 나쁜 사람 손에 들어가면 보이스피싱·명의도용에 바로 쓰입니다.</span></div>
      <div class="mini"><div class="ic">🧱</div><b>두 개의 방</b>
        <span>업무망(회사 안·안전) → 인터넷망(바깥). 파일을 밖으로 보낼 때 개인정보가 새어 나갈 수 있어요.</span></div>
      <div class="mini"><div class="ic">⚖️</div><b>법·국정원 규정</b>
        <span>개인정보보호법과 국정원 N2SF는 “내보낼 땐 가려라”고 정합니다. 한전은 공기업이라 특히 엄격.</span></div>
    </div>
    <div class="callout info big"><b>그래서 —</b> 개인정보가 든 파일을 밖으로 보낼 땐, 이 프로그램으로
      <b>개인정보를 가린 뒤</b> 내보내야 합니다.</div>
    <h3>‘마스킹’이 뭔가요?</h3>
    <p>개인정보를 <b>별표(*)로 가려 못 알아보게</b> 만드는 것입니다. (한 번 가리면 되돌릴 수 없습니다.)</p>
    <table>
      <tr><th>원래 내용</th><th>마스킹 후</th></tr>
      <tr><td>홍길동</td><td><code class="mask">홍*동</code></td></tr>
      <tr><td>010-1234-5678</td><td><code class="mask">010-****-5678</code></td></tr>
      <tr><td>900101-1234567</td><td><code class="mask">******-*******</code></td></tr>
    </table>
  </section>

  <!-- SETUP -->
  <section id="setup">
    <h2><span class="num">2</span> 처음 준비 <span style="font-size:.9rem;color:var(--muted)">(딱 한 번만)</span></h2>
    <div class="steps">
      <div class="step"><span class="k">1</span><b>프로그램을 더블클릭해서 실행</b>합니다. 따로 설치 과정은 없어요.
        <div class="callout warn" style="margin:10px 0 0">⏳ <b>처음 실행은 10~20초</b>쯤 걸립니다. 창이 바로 안 떠도 기다려 주세요.</div></div>
      <div class="step"><span class="k">2</span><b>화면 오른쪽 아래 시계 옆(트레이)</b>에 자물쇠 아이콘(🔒)이 생깁니다.
        <div class="annot"><div class="shotwrap"><img src="%%SHOT_TRAYICON%%" alt="트레이 자물쇠 아이콘"></div>
          <ol class="legend"><li class="key"><span class="ln">1</span><span><b>주황 네모로 표시된 자물쇠 아이콘</b> — 이게 이 프로그램입니다(계속 켜져 있어요). 마우스를 올리면 “엑셀파일 개인정보 마스킹” 툴팁이 뜹니다. <b>안 보이면 ∧(위쪽 화살표)</b>를 눌러 숨은 아이콘을 펼치세요.</span></li></ol></div></div>
      <div class="step"><span class="k">3</span><b>자물쇠를 마우스 오른쪽 버튼으로 클릭</b> → 나오는 메뉴에서
        <b>[⑥ 오른쪽클릭에 ‘마스킹 사본 만들기’ 추가]</b> 를 한 번 누르세요.
        <div class="annot"><div class="shotwrap"><img src="%%SHOT_TRAYMENU%%" alt="트레이 오른쪽클릭 메뉴">
          <span class="hot" style="left:7%;top:10%">1</span>
          <span class="hot" style="left:7%;top:20%">2</span>
          <span class="hot" style="left:7%;top:34%">3</span>
          <span class="hot" style="left:7%;top:44%">4</span>
          <span class="hot" style="left:7%;top:54%">5</span>
          <span class="hot" style="left:7%;top:68%">6</span>
          <span class="hot" style="left:7%;top:78%">7</span>
          <span class="hot" style="left:7%;top:91%">8</span></div>
          <ol class="legend">
            <li><span class="ln">1</span><span><b>마스킹 도구 열기</b> — 파일을 열어 확인하며 마스킹하는 창(방법 2).</span></li>
            <li class="key"><span class="ln">2</span><span><b>📖 사용설명서 열기</b> — 바로 이 설명서를 다시 엽니다.</span></li>
            <li><span class="ln">3</span><span><b>자동 마스킹 폴더 열기</b> — 파일을 넣으면 자동 처리되는 폴더(방법 3).</span></li>
            <li><span class="ln">4</span><span><b>완료된 파일 폴더 열기</b> — 마스킹이 끝난 결과 파일이 모이는 곳.</span></li>
            <li><span class="ln">5</span><span><b>자동 마스킹 잠시 멈춤</b> — 자동 처리를 잠깐 멈춥니다.</span></li>
            <li class="key"><span class="ln">6</span><span><b>오른쪽클릭에 ‘마스킹 사본 만들기’ 추가</b> — <b>처음에 한 번 누르세요.</b> 파일 우클릭 메뉴에 마스킹 항목이 생깁니다.</span></li>
            <li><span class="ln">7</span><span><b>오른쪽클릭 메뉴에서 빼기</b> — 우클릭 메뉴 항목을 제거합니다.</span></li>
            <li><span class="ln">8</span><span><b>종료</b> — 프로그램을 완전히 끕니다.</span></li>
          </ol></div></div>
    </div>
    <div class="callout ok">✅ “추가되었습니다” 알림이 뜨면 준비 끝입니다!</div>
  </section>

  <!-- DRM 사전작업 -->
  <section id="drm" style="border:2px solid #fca5a5">
    <h2><span class="num" style="background:#dc2626">3</span> 먼저! DRM 해제 <span style="font-size:.9rem;color:var(--muted)">(꼭 필요한 사전 작업)</span></h2>
    <div class="callout warn big"><b>⚠️ 마스킹보다 먼저 하세요.</b> 사내 파일은 보안 정책(DRM)으로 <b>잠겨(암호화)</b> 있어,
      잠긴 채로는 마스킹도 반출도 되지 않습니다. <b>DRM을 먼저 해제</b>한 뒤 마스킹하세요.</div>
    <div class="steps">
      <div class="step"><span class="k">1</span>대상 <b>엑셀 파일을 마우스 오른쪽 버튼</b>으로 클릭합니다.</div>
      <div class="step"><span class="k">2</span>메뉴에서 <b>[DRM 설정 메뉴]</b>에 마우스를 올립니다. → 오른쪽에 작은 메뉴가 펼쳐집니다.</div>
      <div class="step"><span class="k">3</span><span class="clickrow">펼쳐진 메뉴에서 <b>[암호화 해제(부서장)]</b>를 클릭합니다.<span class="lottie sm" data-anim="click"></span></span>
        <figure class="shot"><img src="%%IMG_DRM%%" alt="DRM 해제 방법: 우클릭 → DRM 설정 메뉴 → 암호화 해제(부서장)">
          <figcaption><span class="tag-real">실제 화면</span> 파일 우클릭 → <b>DRM 설정 메뉴</b> → <b>암호화 해제(부서장)</b> 클릭</figcaption></figure></div>
    </div>
    <div class="callout ok big">✅ DRM이 풀렸으면 이제 <b>아래 마스킹</b>을 진행하세요. (잠금이 남아 있으면 “파일을 열 수 없습니다”가 뜹니다.)</div>
  </section>

  <!-- USE -->
  <section id="use">
    <h2><span class="num">4</span> 가장 쉬운 사용법 — 오른쪽 버튼 3단계 ⭐</h2>
    <div class="sub">개인정보가 든 엑셀을 인터넷망으로 보내야 할 때.</div>
    <div class="steps">
      <div class="step"><span class="k">1</span><b>파일 위에서 마우스 오른쪽 버튼</b>을 누릅니다. &nbsp;📄 직원명부.xlsx</div>
      <div class="step"><span class="k">2</span><span class="clickrow">메뉴에서 <b>[🔒 개인정보 마스킹 사본 만들기]</b> 를 누릅니다.<span class="lottie sm" data-anim="click"></span></span>
        <div class="annot"><div class="shotwrap"><img src="%%SHOT_RIGHTCLICK%%" alt="엑셀 파일 오른쪽클릭 메뉴">
          <span class="hot" style="left:40%;top:18%">1</span></div>
          <ol class="legend"><li class="key"><span class="ln">1</span><span><b>🔒 개인정보 마스킹 사본 만들기</b> — 이 항목을 클릭하면 개인정보를 가린 사본이 원본 옆에 새로 만들어집니다.</span></li></ol></div></div>
      <div class="step"><span class="k">3</span>안내창이 뜨고, <b>같은 폴더에 “_마스킹” 파일이 새로</b> 생깁니다.
        <div class="win"><div class="bar"><span class="dot"></span>✅ 완료 안내창</div>
          <div style="padding:16px">
            <div class="clickrow"><span class="lottie" data-anim="check"></span><b style="font-size:1.05rem">마스킹 사본 생성 완료</b></div>
            개인정보를 가린 사본을 만들었습니다.<br><br>
            만들어진 파일: <b>직원명부_마스킹.xlsx</b>
            <span style="color:var(--muted)">(원본과 같은 폴더에 새로 생겼어요)</span><br><br>
            가린 개인정보 (총 12건): 이름×5 · 전화번호×4 · 주민등록번호×3<br><br>
            ▶ FTC로는 이 ‘_마스킹’ 파일을 내보내세요.
          </div></div>
        <div class="files">
          <div class="row">📄 직원명부.xlsx <span class="badge no">원본 · 내보내지 마세요</span></div>
          <div class="row">📄 직원명부_마스킹.xlsx <span class="badge ok">이걸 FTC로 내보내세요</span></div>
        </div></div>
    </div>
    <div class="callout ok big">끝입니다! FTC로는 반드시 <b>“_마스킹”이 붙은 파일</b>을 내보내세요.</div>
  </section>

  <!-- WINDOW -->
  <section id="window">
    <h2><span class="num">5</span> 도구 창으로 하기 <span style="font-size:.9rem;color:var(--muted)">(차근차근 확인하며)</span></h2>
    <div class="sub">트레이 자물쇠를 더블클릭하면 아래 창이 열립니다. 파일을 열면 <b>가릴 칸이 자동으로 체크</b>됩니다.</div>
    <figure class="shot">
      <img src="%%IMG_MAIN%%" alt="마스킹 도구 창">
      <figcaption><span class="tag-real">실제 화면</span> ① [파일 열기]로 엑셀 선택 → ② 가릴 칸이 자동 체크·등급(C/S/O) 표시 → ③ 확인 후 [마스킹 실행 및 저장]</figcaption>
    </figure>
    <ul class="plain">
      <li><b>등급 색</b> — <span style="color:#dc2626;font-weight:700">C(빨강)</span> 완전 가림 · <span style="color:#f59e0b;font-weight:700">S(주황)</span> 부분 가림 · <span style="color:#94a3b8;font-weight:700">O(회색)</span> 그대로 둠</li>
      <li><b>오탐 %</b> — “이 칸을 잘못 가릴 가능성”의 추정치입니다. 근거도 함께 보여줍니다.</li>
      <li><b>체크박스</b> — 체크된 칸만 가려집니다. 직접 켜고 끌 수 있어요.</li>
    </ul>
  </section>

  <!-- FOLDER -->
  <section id="folder">
    <h2><span class="num">6</span> 폴더에 넣기만 하면 자동 처리</h2>
    <div class="sub">트레이 메뉴 → <b>[자동 마스킹 폴더 열기]</b> 를 누르면 아래 폴더가 열립니다.
      여기에 파일을 <b>끌어다 넣기만</b> 하면 자동으로 마스킹됩니다.</div>
    <div class="annot"><div class="shotwrap"><img src="%%SHOT_FOLDERS%%" alt="자동 마스킹 폴더 구조">
      <span class="hot" style="left:40%;top:16%">1</span></div>
      <ol class="legend"><li class="key"><span class="ln">1</span><span>여기가 <b>자동_마스킹_폴더</b>(문서 폴더 안). 이 안에 아래 세 폴더가 자동으로 만들어집니다. 각 폴더의 역할은 아래 표를 보세요.</span></li></ol></div>
    <table>
      <tr><th>폴더</th><th>무슨 폴더인가요?</th></tr>
      <tr><td><b>자동_마스킹_폴더</b></td><td>여기에 파일을 <b>넣으면 자동으로 처리 시작</b>. (넣는 곳)</td></tr>
      <tr><td><b>_완료(마스킹된파일)</b></td><td>마스킹이 끝난 <b>결과 파일</b>이 생기는 곳. → <b>FTC로는 여기 파일을 내보내세요.</b></td></tr>
      <tr><td><b>_원본보관</b></td><td>처리한 <b>원본</b>이 자동으로 옮겨져 보관됩니다. (실수로 원본을 또 보내지 않도록)</td></tr>
      <tr><td><b>_오류</b></td><td>처리하지 못한 파일이 모입니다. 보통 <b>파일이 Excel에서 열려 있거나</b> 손상된 경우예요.</td></tr>
    </table>
    <div class="callout info">여러 파일을 한꺼번에 넣어도 됩니다. 완료되면 오른쪽 아래에 “마스킹 완료” 알림이 뜹니다.</div>
  </section>

  <!-- CHECK -->
  <section id="check">
    <h2><span class="num">7</span> 보내기 전, 꼭 확인하세요</h2>
    <div class="sub">자동이라 가끔 놓칠 수 있어요. 아래를 눌러 확인하며 체크하세요. (체크는 저장됩니다)</div>
    <div class="progress"><i id="bar"></i></div>
    <div id="pct" style="font-size:.86rem;color:var(--muted);margin-bottom:8px">0 / 3 확인함</div>
    <ul class="check" id="checklist">
      <li><input type="checkbox"><span>파일 이름에 <b>‘_마스킹’</b> 이 붙어 있는지 확인했다. (원본을 실수로 보내면 안 됨)</span></li>
      <li><input type="checkbox"><span>파일을 열어 <b>개인정보가 별표(*)로 가려졌는지</b> 눈으로 확인했다.</span></li>
      <li><input type="checkbox"><span>엑셀 아래 <b>[분류리포트]</b> 시트에서 무엇을 왜 가렸는지 확인했다.</span></li>
    </ul>
    <div class="callout info">엑셀 아래쪽 시트 탭: [ 마스킹결과 ] &nbsp;<b>[ 분류리포트 ]</b>&nbsp; [ N2SF근거 ] — 가운데를 누르면 목록이 나옵니다.</div>
  </section>

  <!-- RULE -->
  <section id="rule">
    <h2><span class="num">8</span> 안 가려진 칸이 있을 때 — “항상 가릴 칸” 정하기</h2>
    <p>이 프로그램은 <b>글자 모양을 보고 스스로</b> 이름·주소·전화번호를 찾습니다. (칸 제목을 몰라도 됩니다.)
       하지만 <b>우리 회사에만 있는 특별한 칸</b>(예: 계약번호·상호)은 못 알아챌 수 있어요.
       그럴 땐 <b>한 번만 알려주면 다음부터 기억</b>합니다.</p>
    <p>도구 창 위쪽 <b>[📌 항상 마스킹할 칸 정하기]</b> → <b>칸</b>과 <b>가리는 방법</b>을 고르고 <b>[＋ 추가]</b>.</p>
    <figure class="shot">
      <img src="%%IMG_RULES%%" alt="항상 마스킹할 칸 정하기 창">
      <figcaption><span class="tag-real">실제 화면</span> ‘계약번호’ 칸을 ‘숫자만 가림’으로 정해 둔 모습. 이제부터 그 칸은 항상 자동으로 가려집니다.</figcaption>
    </figure>
    <div class="callout ok">한 번만 정하면 계속 기억해요. 우클릭·자동 폴더에서도 똑같이 적용됩니다.</div>
  </section>

  <!-- FAQ -->
  <section id="faq">
    <h2><span class="num">9</span> 자주 묻는 질문</h2>
    <div class="faq">
      <details><summary>어려운 프로그램을 따로 깔아야 하나요?</summary>
        <div class="ans">아니요. 받은 파일 하나만 실행하면 됩니다.</div></details>
      <details><summary>원본 파일이 없어지거나 바뀌나요?</summary>
        <div class="ans">아니요. 원본은 그대로 있고, 옆에 “_마스킹” 사본이 새로 생깁니다.</div></details>
      <details><summary>오른쪽 버튼 메뉴에 항목이 안 보여요.</summary>
        <div class="ans">위 <b>2. 처음 준비</b>의 3단계(자물쇠 → ‘마스킹 사본 만들기’ 추가)를 한 번 해 주세요.</div></details>
      <details><summary>개인정보가 아닌데 가려졌어요 / 가려야 하는데 안 가려졌어요.</summary>
        <div class="ans">자동이라 가끔 그럴 수 있어요. <b>7번(항상 가릴 칸 정하기)</b> 이나 <b>4번(도구 창에서 체크 조정)</b>으로 고칠 수 있습니다.</div></details>
      <details><summary>처음 실행이 느려요. 고장인가요?</summary>
        <div class="ans">정상입니다. 처음만 조금 느리고, 두 번째부터 빨라집니다.</div></details>
      <details><summary>내가 언제 뭘 했는지 기록이 남나요?</summary>
        <div class="ans">“누가·언제·어떤 파일을·몇 건 가렸는지”만 남습니다. 파일 속 실제 개인정보 내용은 절대 기록하지 않습니다.</div></details>
      <details><summary>어떤 파일을 넣을 수 있나요?</summary>
        <div class="ans">엑셀(.xlsx, .xls)과 CSV(.csv). 한전 시스템에서 내려받은 엑셀도 대부분 됩니다.</div></details>
    </div>
  </section>

  <!-- WORDS -->
  <section id="words">
    <h2><span class="num">10</span> 용어 풀이</h2>
    <table>
      <tr><th>말</th><th>쉬운 뜻</th></tr>
      <tr><td>마스킹</td><td>개인정보를 별표(*)로 가려 못 알아보게 하는 것 (되돌릴 수 없음)</td></tr>
      <tr><td>사본</td><td>원본을 복사해서 새로 만든 파일</td></tr>
      <tr><td>FTC</td><td>회사 안(업무망)에서 바깥(인터넷망)으로 파일을 보내는 프로그램</td></tr>
      <tr><td>업무망 / 인터넷망</td><td>회사 안쪽 방 / 바깥과 연결된 방</td></tr>
      <tr><td>트레이</td><td>화면 오른쪽 아래 시계 옆, 작은 아이콘들이 모인 곳</td></tr>
      <tr><td>칸(열)</td><td>엑셀에서 세로 줄. 예: ‘이름’ 칸, ‘전화번호’ 칸</td></tr>
    </table>
    <div class="callout warn">파일이 안 열리면: 엑셀을 열어 <b>[다른 이름으로 저장 → Excel 통합 문서(.xlsx)]</b>로 다시 저장 후 시도하거나,
      <b>사내 AI 담당자</b>에게 문의하세요.</div>
  </section>

  <footer>이 프로그램은 개인정보보호법과 국정원 N2SF 정책에 따라, 업무망 → 인터넷망 파일 반출 시
    개인정보를 안전하게 지키기 위해 만들어졌습니다.</footer>
</div>

<script>%%LOTTIE_PLAYER%%</script>
<script>
  window.LOTTIE_ANIMS = %%LOTTIE_ANIMS%%;
  window.addEventListener('load', function(){
    if(!window.lottie) return;
    document.querySelectorAll('.lottie').forEach(function(el){
      var a = window.LOTTIE_ANIMS[el.getAttribute('data-anim')];
      if(!a) return;
      lottie.loadAnimation({container:el, renderer:'svg',
        loop:(el.getAttribute('data-loop')!=='0'), autoplay:true,
        animationData: JSON.parse(JSON.stringify(a))});
    });
  });
</script>
<script>
  (function(){
    var list=document.getElementById('checklist'); if(!list)return;
    var items=list.querySelectorAll('li'), boxes=list.querySelectorAll('input');
    var bar=document.getElementById('bar'), pct=document.getElementById('pct');
    function key(i){return 'ftcmask_check_'+i;}
    function refresh(){
      var done=0;
      boxes.forEach(function(b,i){ items[i].classList.toggle('done',b.checked); if(b.checked)done++; });
      bar.style.width=(done/boxes.length*100)+'%';
      pct.textContent=done+' / '+boxes.length+' 확인함';
    }
    boxes.forEach(function(b,i){
      try{ b.checked=localStorage.getItem(key(i))==='1'; }catch(e){}
      items[i].addEventListener('click',function(ev){
        if(ev.target!==b){ b.checked=!b.checked; }
        try{ localStorage.setItem(key(i), b.checked?'1':'0'); }catch(e){}
        refresh();
      });
    });
    refresh();
  })();
</script>
</body>
</html>
"""


def build(out_dir=None, open_after=False):
    """단일 HTML 생성 → 경로 반환. out_dir 미지정 시 데이터 폴더."""
    out_dir = out_dir or _data_dir()
    os.makedirs(out_dir, exist_ok=True)
    import base64 as _b64
    player_js = _b64.b64decode(PLAYER_B64).decode("utf-8") if PLAYER_B64 else ""
    anims_js = "{" + ",".join(f'"{k}":{v}' for k, v in ANIM.items()) + "}"
    html = (HTML
            .replace("%%IMG_MAIN%%", IMAGES.get("main", ""))
            .replace("%%IMG_RULES%%", IMAGES.get("rules", ""))
            .replace("%%IMG_DRM%%", STATIC.get("drm", ""))
            .replace("%%LOTTIE_PLAYER%%", player_js)
            .replace("%%LOTTIE_ANIMS%%", anims_js)
            .replace("%%SHOT_TRAYICON%%", SHOTS.get("tray_icon", ""))
            .replace("%%SHOT_TRAYMENU%%", SHOTS.get("tray_menu", ""))
            .replace("%%SHOT_RIGHTCLICK%%", SHOTS.get("rightclick", ""))
            .replace("%%SHOT_FOLDERS%%", SHOTS.get("folders", "")))
    path = os.path.join(out_dir, "사용설명서.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    if open_after:
        webbrowser.open("file:///" + path.replace("\\", "/"))
    return path


def open_manual():
    """앱에서 호출: HTML을 새로 만들고 기본 브라우저로 연다(매일 삭제돼도 매번 재생성)."""
    return build(open_after=True)


if __name__ == "__main__":
    p = build(out_dir=os.path.join(os.path.dirname(os.path.abspath(__file__)), "docs"),
              open_after=("--open" in sys.argv))
    print("생성됨:", p)
