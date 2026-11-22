/* Rasta House authenticated demo wallet.
 * Session tokens authenticate the player; game outcomes still require
 * server-authoritative round validation before any real-money use.
 */
(function(){
  'use strict';
  const ENDPOINT='https://script.google.com/macros/s/AKfycby5oeX598AtqVX_owDmR1yKiAQ6z0fjquHzf4YJccbyBqF1GwrGnRwyFt_pVvGyGl7j/exec';
  function user(){try{return JSON.parse(localStorage.getItem('rasta_user')||'null')}catch(e){return null}}
  function save(balance){
    const u=user();if(!u)return;
    u.balance=Number(balance);
    localStorage.setItem('rasta_user',JSON.stringify(u));
    localStorage.setItem('rasta_demo_credits',String(u.balance));
  }
  async function request(action,extra){
    const u=user();
    if(!u||!u.sessionToken)throw new Error('Your session has expired. Please log in again.');
    const params=new URLSearchParams(Object.assign({action:action,token:u.sessionToken},extra||{}));
    const response=await fetch(ENDPOINT,{method:'POST',mode:'cors',cache:'no-store',headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},body:params});
    if(!response.ok)throw new Error('Wallet server is unavailable.');
    const data=await response.json();
    if(!data||data.status!=='success')throw new Error(data&&data.message||'Wallet request failed.');
    if(data.balance!==undefined)save(data.balance);
    return data;
  }
  window.RastaWallet={
    endpoint:ENDPOINT,
    request:request,
    sync:async function(){const d=await request('getBalance');let balance=Number(d.balance);try{const pending=JSON.parse(localStorage.getItem('rasta_pending_game_payout')||'null');if(pending&&pending.roundId&&Number(pending.amount)>0){balance=await this.credit(pending.amount,pending.roundId);localStorage.removeItem('rasta_pending_game_payout')}}catch(e){console.warn('Pending game payout will retry on next sync.',e)}return balance},
    debit:async function(amount,roundId){const d=await request('gameDebit',{roundId:String(roundId||''),amount:String(amount||0)});return Number(d.balance)},
    credit:async function(amount,roundId){const d=await request('gameCredit',{roundId:String(roundId||''),amount:String(amount||0)});return Number(d.balance)},
    logout:async function(){try{await request('logout')}catch(e){}}
  };
})();
