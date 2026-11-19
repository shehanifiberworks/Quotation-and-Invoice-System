/* SFG Cloud connector. The user-provided public deployment endpoint is prefilled; no access key is shipped.
   Writes are verified by a readback because Apps Script cross-origin POST responses
   may be opaque. The frontend never reports an unverified cloud save as success. */
(function (global) {
  'use strict';
  const CONFIG = 'sfg_cloud_config_v1';
  // User-provided Web App URL. This is a public endpoint, not an authentication secret.
  const DEFAULT_URL = 'https://script.google.com/macros/s/AKfycbwGBoTdPg17xUUntUhxsKKPPdLR9LPKeU6oBnTRHuJHscQECJP2AwA7Ke64RNhqHeByRg/exec';
  const urlOK = value => /^https:\/\/script\.google\.com\/macros\/s\/[A-Za-z0-9_-]+\/exec(?:\?.*)?$/.test(String(value || '').trim());
  function config() { try { const c = JSON.parse(localStorage.getItem(CONFIG) || '{}'); return c && typeof c === 'object' ? {...c, url: c.url || DEFAULT_URL} : {url: DEFAULT_URL}; } catch (_) { return {url: DEFAULT_URL}; } }
  function setConfig(url, accessKey) {
    const clean = String(url || '').trim().split('?')[0];
    const key = String(accessKey || '').trim();
    if (!urlOK(clean)) throw Error('Enter a deployed Google Apps Script Web App URL ending in /exec.');
    if (key.length < 16) throw Error('Enter the access key configured in Apps Script (at least 16 characters).');
    localStorage.setItem(CONFIG, JSON.stringify({url:clean, accessKey:key}));
  }
  function enabled() { const c=config();return urlOK(c.url)&&String(c.accessKey||'').length>=16; }
  let counter = 0;
  function get(action, other={}, timeout=15000) {
    if (!enabled()) return Promise.reject(Error('Google Sheets connection is not configured.'));
    const c = config(), callback = 'sfgCloudCB_' + (++counter) + '_' + Date.now();
    return new Promise((resolve,reject)=>{
      const params = new URLSearchParams({action, token:c.accessKey, callback,...other});
      const script=document.createElement('script');let done=false;
      const cleanup=()=>{if(done)return;done=true;clearTimeout(timer);delete global[callback];script.remove();};
      global[callback]=response=>{cleanup();if(!response || response.ok===false || response.error)reject(Error(response?.error||'Cloud request failed'));else resolve(response);};
      script.onerror=()=>{cleanup();reject(Error('Google Sheets endpoint unavailable. Check URL, deployment access and internet.'));};
      const timer=setTimeout(()=>{cleanup();reject(Error('Cloud response timed out.'));},timeout);
      script.src = c.url + '?' + params.toString();
      document.head.appendChild(script);
    });
  }
  function requestId() { return 'sfg_' + Date.now() + '_' + Math.random().toString(36).slice(2,10); }
  async function write(action,key,data) {
    if (!enabled()) return {ok:false, configured:false, error:'Cloud not configured'};
    const id=requestId(),c=config();
    try {
      // Apps Script may save successfully yet the browser reject its redirect.
      // Always attempt an independent receipt check, even if fetch throws.
      let postIssue='';
      try {
        await fetch(c.url, {method:'POST',mode:'no-cors',credentials:'omit',
          headers:{'Content-Type':'text/plain;charset=utf-8'},
          body:JSON.stringify({action,key,data,token:c.accessKey,requestId:id})});
      } catch(e) {postIssue=e.message||String(e);}
      let lastError=postIssue||'Cloud save was not confirmed.';
      for(const delay of [250,500,900,1600,2600]) {
        await new Promise(r=>setTimeout(r,delay));
        try {
          const result = await get('verify',{requestId:id,key:String(key||'')},10000);
          if(result.done) return {ok:true,configured:true};
          if(result.rejected) return {ok:false,configured:true,error:result.error||'Cloud rejected the change.'};
        } catch(e) { lastError=e.message; }
      }
      return {ok:false,configured:true,error:lastError};
    } catch(e) { return {ok:false,configured:true,error:e.message||String(e)}; }
  }
  async function list() {const response=await get('list');return Array.isArray(response.documents)?response.documents:[];}
  async function test() {const response=await get('ping');return !!response.ok;}
  async function nextNumber() {const response=await get('nextNumber');return response.nextNumber;}
  global.SFGCloud={config,setConfig,enabled,list,test,nextNumber,write};
})(window);
