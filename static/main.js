(function(){
  const timerEl = document.getElementById('timer');
  if (!timerEl || !window.TODAY_SCHEDULE) return;
  const subjectEl = document.getElementById('current-subject');
  const classEl = document.getElementById('current-class');
  const periodEl = document.getElementById('current-period');
  const startEl = document.getElementById('start');
  const endEl = document.getElementById('end');
  const bell = document.getElementById('bell');
  const soundToggle = document.getElementById('sound-toggle');
  function hhmmToDate(hhmm){ const [h,m]=hhmm.split(':').map(Number); const n=new Date(); return new Date(n.getFullYear(),n.getMonth(),n.getDate(),h,m,0); }
  function findCurrent(){ const now=new Date(); for(const p of window.TODAY_SCHEDULE){ const st=hhmmToDate(p.start_time); const en=hhmmToDate(p.end_time); if(now>=st && now<=en) return {...p,st,en}; } for(const p of window.TODAY_SCHEDULE){ const st=hhmmToDate(p.start_time); if(now<st) return {...p,st,en:hhmmToDate(p.end_time)}; } return null; }
  function fmt(ms){ const sec=Math.max(0,Math.floor(ms/1000)); const h=String(Math.floor(sec/3600)).padStart(2,'0'); const m=String(Math.floor((sec%3600)/60)).padStart(2,'0'); const s=String(sec%60).padStart(2,'0'); return `${h}:${m}:${s}`; }
  let lastPeriodId=null; setInterval(()=>{ const cur=findCurrent(); if(!cur){ timerEl.textContent='--:--:--'; timerEl.classList.remove('warn','danger'); subjectEl.textContent='انتهى جدول اليوم'; classEl.textContent='—'; periodEl.textContent='—'; startEl.textContent='--:--'; endEl.textContent='--:--'; return; } if(lastPeriodId!==cur.id){ subjectEl.textContent=cur.subject; classEl.textContent=cur.class_name; periodEl.textContent=cur.period; startEl.textContent=cur.start_time; endEl.textContent=cur.end_time; lastPeriodId=cur.id; timerEl.classList.remove('warn','danger'); } const now=new Date(); const remaining=cur.en-now; timerEl.textContent=fmt(remaining); const ten=10*60*1000, five=5*60*1000; timerEl.classList.remove('warn','danger'); if(remaining<=ten && remaining>five) timerEl.classList.add('warn'); else if(remaining<=five) timerEl.classList.add('danger'); if(remaining<=0){ if(soundToggle && soundToggle.checked && bell){ bell.currentTime=0; bell.play().catch(()=>{}); } }
  },1000);
})();
