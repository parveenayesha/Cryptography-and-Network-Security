(function () {
  'use strict';
  const menu = document.querySelector('.menu-toggle');
  const nav = document.querySelector('#nav');
  if (menu && nav) {
    menu.addEventListener('click', () => {
      const open = nav.classList.toggle('open');
      menu.setAttribute('aria-expanded', String(open));
    });
    nav.querySelectorAll('a').forEach(link => link.addEventListener('click', () => {
      nav.classList.remove('open'); menu.setAttribute('aria-expanded', 'false');
    }));
  }

  function caesar(text, shift, decrypt) {
    const amount = ((Number(shift) % 26) + 26) % 26 * (decrypt ? -1 : 1);
    return String(text).replace(/[a-z]/gi, ch => {
      const base = ch >= 'a' && ch <= 'z' ? 97 : 65;
      return String.fromCharCode(((ch.charCodeAt(0) - base + amount + 26) % 26) + base);
    });
  }

  const input = document.querySelector('#plainText');
  const shift = document.querySelector('#shift');
  const mode = document.querySelector('#mode');
  const output = document.querySelector('#cipherResult');
  const run = document.querySelector('#runCipher');
  function runCipher() {
    const key = Math.max(0, Math.min(25, Number.parseInt(shift.value || '0', 10)));
    shift.value = String(key);
    output.textContent = caesar(input.value, key, mode.value === 'decrypt');
  }
  if (run) run.addEventListener('click', runCipher);
  if (input) input.addEventListener('input', runCipher);
  if (shift) shift.addEventListener('input', runCipher);
  if (mode) mode.addEventListener('change', runCipher);
  runCipher();

  const quiz = document.querySelector('#quizForm');
  const quizResult = document.querySelector('#quizResult');
  if (quiz) quiz.addEventListener('submit', event => {
    event.preventDefault();
    const answers = { q1: 'b', q2: 'a', q3: 'a' };
    let score = 0;
    const formData = new FormData(quiz);
    for (const [question, answer] of Object.entries(answers)) {
      if (formData.get(question) === answer) score += 1;
    }
    const feedback = score === 3 ? 'Excellent: 3/3. You can explain the key protections.' :
      score === 2 ? 'Good: 2/3. Review the feedback and try again.' :
      'Score: ' + score + '/3. Revisit the unit notes and try again.';
    quizResult.textContent = feedback;
  });
})();
