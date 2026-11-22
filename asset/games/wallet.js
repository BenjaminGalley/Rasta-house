/* Rasta House server-authoritative demo wallet */
(function(){
  'use strict';
  const ENDPOINT='https://script.google.com/macros/s/AKfycby5oeX598AtqVX_owDmR1yKiAQ6z0fjquHzf4YJccbyBqF1GwrGnRwyFt_pVvGyGl7j/exec';
  function user(){try{return JSON.parse(localStorage.getItem('rasta_user')||'null')}catch(e){return null}}
  function save(balance){const u=user();if(!u)return;u.balance=Number(balance);localStorage.setItem('rasta_user',JSON.stringify(u));localStorage.setItem('rasta_demo_credits',String(u.balance))}
  async function call(action,roundId,amount){
    const u=user();if(!u||!u.phone)throw new Error('Please log in first.');
    const q=new URLSearchParams({action,phone:String(u.phone),roundId:String(roundId||''),amount:String(amount||0),_ :String(Date.now())});
    const response=await fetch(ENDPOINT+'?'+q.toString(),{cache:'no-store'});
    const data=await response.json();
    if(!data||data.status!=='success')throw new Error(data&&data.message||'Wallet connection failed.');
    save(data.balance);return Number(data.balance);
  }
  window.RastaWallet={
    endpoint:ENDPOINT,
    sync:async function(){const u=user();if(!u||!u.phone)throw new Error('Please log in first.');const r=await fetch(ENDPOINT+'?action=getBalance&phone='+encodeURIComponent(u.phone)+'&_='+Date.now(),{cache:'no-store'});const d=await r.json();if(d.status!=='success')throw new Error(d.message||'Could not load balance.');save(d.balance);return Number(d.balance)},
    debit:function(amount,roundId){return call('gameDebit',roundId,amount)},
    credit:function(amount,roundId){return call('gameCredit',roundId,amount)}
  };
})();