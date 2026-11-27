---
layout: single
title: 예약 자동화 코드 만들기
permalink: /generator/
excerpt: 가게 이름, 정원, 폼 질문만 넣으면 구글폼 예약 확정·대기 메일 코드를 바로 만들어 줍니다. 무료, 설치 없이 브라우저에서.
---

구글폼으로 예약을 받는 가게용입니다. 아래 칸을 채우면 **우리 가게 값이 들어간 앱스스크립트 코드**가 나옵니다. 코드를 고칠 필요 없이 그대로 붙여 넣으면 됩니다.

- 입력한 내용은 이 브라우저 안에서만 쓰이고 어디로도 전송되지 않습니다.
- 붙여 넣는 방법은 [예약 확정·대기 메일 레시피]({{ '/booking-form-auto-confirm/' | relative_url }})의 2단계부터 따라 하세요.

<style>
.gen label{display:block;font-weight:600;margin:14px 0 4px}.gen input[type=text],.gen input[type=number],.gen input[type=email]{width:100%;max-width:420px;padding:8px 10px;border:1px solid #cbd5e1;border-radius:6px;font-size:16px}
.gen fieldset{border:1px solid #e2e8f0;border-radius:10px;padding:8px 16px 16px;margin:18px 0}.gen legend{font-weight:700;padding:0 6px}
.gen .row{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:0 16px}.gen small{color:#64748b}
.gen button{margin-top:14px;padding:10px 18px;border:0;border-radius:8px;background:#12306B;color:#fff;font-size:16px;font-weight:700;cursor:pointer}
.gen pre{max-height:420px;overflow:auto;background:#0f172a;color:#e2e8f0;padding:14px;border-radius:10px;font-size:13px;white-space:pre}
.gen .msg{color:#b91c1c;font-weight:600;margin-top:10px}
</style>

<div class="gen">
<fieldset><legend>1. 가게 정보</legend>
<label for="g-shop">가게 이름 <small>(메일 제목과 본문에 들어갑니다)</small></label><input type="text" id="g-shop" placeholder="예: 몽이공방">
<label for="g-owner">새 예약 알림을 받을 사장님 메일</label><input type="email" id="g-owner" placeholder="예: boss@example.com">
<div class="row">
<div><label for="g-cap">같은 시간에 받을 최대 인원</label><input type="number" id="g-cap" min="1" value="4"></div>
<div><label for="g-sheet">정리용 시트 이름</label><input type="text" id="g-sheet" value="예약목록"></div>
</div>
</fieldset>

<fieldset><legend>2. 내 폼의 질문 제목</legend>
<small>폼에 적은 질문 제목과 글자 하나까지 같아야 합니다. 그대로 쓰고 있다면 바꾸지 않아도 됩니다.</small>
<div class="row">
<div><label for="g-f-name">이름</label><input type="text" id="g-f-name" value="이름"></div>
<div><label for="g-f-phone">연락처</label><input type="text" id="g-f-phone" value="연락처"></div>
<div><label for="g-f-email">이메일</label><input type="text" id="g-f-email" value="이메일"></div>
<div><label for="g-f-date">날짜</label><input type="text" id="g-f-date" value="희망 날짜"></div>
<div><label for="g-f-time">시간</label><input type="text" id="g-f-time" value="희망 시간"></div>
<div><label for="g-f-people">인원</label><input type="text" id="g-f-people" value="인원"></div>
</div>
</fieldset>

<fieldset><legend>3. 캘린더</legend>
<label><input type="checkbox" id="g-cal"> 확정된 예약을 구글 캘린더에도 자동으로 넣기</label>
<div id="g-cal-opt" style="display:none"><label for="g-slot">한 타임 길이(분)</label><input type="number" id="g-slot" min="10" value="120"></div>
</fieldset>

<button type="button" id="g-make">코드 만들기</button>
<div class="msg" id="g-msg" role="alert"></div>

<div id="g-out" style="display:none">
<h2>만들어진 코드</h2>
<button type="button" id="g-copy">코드 복사</button> <small id="g-copied"></small>
<pre id="g-code"></pre>
<p>다음 순서: 응답 시트에서 <strong>확장 프로그램 → Apps Script</strong>를 열고, 있던 내용을 지운 뒤 붙여 넣으세요. 그다음 <code>setupTrigger</code>를 한 번 실행하면 됩니다. 화면은 <a href="{{ '/booking-form-auto-confirm/' | relative_url }}">레시피 글</a>에 있습니다.</p>
</div>
</div>

{% capture tpl_a %}{% include code/booking-auto-confirm.gs %}{% endcapture %}
{% capture tpl_b %}{% include code/booking-with-calendar.gs %}{% endcapture %}
<script src="{{ '/assets/js/booking-generator.js' | relative_url }}"></script>
<script>
(function () {
  var TPL = { a: {{ tpl_a | jsonify }}, b: {{ tpl_b | jsonify }} };
  var $ = function (id) { return document.getElementById(id); };
  $('g-cal').addEventListener('change', function () { $('g-cal-opt').style.display = this.checked ? '' : 'none'; });
  $('g-make').addEventListener('click', function () {
    var msg = $('g-msg'); msg.textContent = '';
    var shop = $('g-shop').value.trim(), owner = $('g-owner').value.trim();
    if (!shop) { msg.textContent = '가게 이름을 넣어 주세요.'; return; }
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(owner)) { msg.textContent = '사장님 메일 주소를 확인해 주세요.'; return; }
    try {
      var useCal = $('g-cal').checked;
      var code = buildBookingCode(useCal ? TPL.b : TPL.a, {
        shopName: shop, ownerEmail: owner, sheetName: $('g-sheet').value.trim(), capacity: $('g-cap').value,
        slotMinutes: useCal ? $('g-slot').value : null,
        fields: { name: $('g-f-name').value.trim(), phone: $('g-f-phone').value.trim(), email: $('g-f-email').value.trim(),
                  date: $('g-f-date').value.trim(), time: $('g-f-time').value.trim(), people: $('g-f-people').value.trim() }
      });
      $('g-code').textContent = code; $('g-out').style.display = ''; $('g-copied').textContent = '';
      $('g-out').scrollIntoView({ behavior: 'smooth' });
    } catch (e) { msg.textContent = e.message; }
  });
  $('g-copy').addEventListener('click', function () {
    var t = $('g-code').textContent;
    (navigator.clipboard ? navigator.clipboard.writeText(t) : Promise.reject()).then(function () { $('g-copied').textContent = '복사했어요.'; },
      function () { var r = document.createRange(); r.selectNodeContents($('g-code')); var s = getSelection(); s.removeAllRanges(); s.addRange(r); $('g-copied').textContent = '코드를 선택했어요. Ctrl+C로 복사하세요.'; });
  });
})();
</script>

---

**이 도구에 대해.** 도구와 코드는 AI(Claude)가 만들었습니다. 만들어진 코드가 입력값대로 동작하는지는 구글 서비스를 흉내 낸 테스트 18개로 확인했습니다(가게 이름·정원·질문 제목 반영, 정원 초과 시 대기 처리, 따옴표가 든 이름 처리 등). 바탕이 되는 코드는 2026년 11월 26일에 실제 구글 계정에서 예약 확정·대기까지 돌려 봤습니다. 캘린더 등록 부분은 아직 실제 계정에서 돌린 결과가 아닙니다. 쓰시기 전에 직접 한 번 예약해 보세요.

직접 설치하기 번거로우시면 크몽 [구글폼·구글시트 자동화 서비스](https://kmong.com/gig/877282)(1만 원부터)에서 맞춰 드려요.
