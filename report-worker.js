(function (root) {
  'use strict';
  if (typeof importScripts === 'function' && !root.SurveyCore) importScripts('survey-core.js');
  function handle(message) {
    const data=message&&message.data?message.data:message; const core=root.SurveyCore; const sections=(data.questions||[]).map((q,i)=>{const out=core.reportSection(data.rows||[],q.display||q.column,q.column,data.breakdownColumns||[],{...(data.options||{}),...(q.options||{}),answerOrder:q.answerOrder||[]}); if(root.postMessage) root.postMessage({id:data.id,progress:{completed:i+1,total:(data.questions||[]).length}}); return out;}); return {id:data.id,sections};
  }
  if (typeof root.addEventListener==='function') root.addEventListener('message',e=>{try{root.postMessage(handle(e));}catch(error){root.postMessage({id:e.data&&e.data.id,error:String(error&&error.message||error)});}});
  if (typeof module==='object'&&module.exports) module.exports={handle};
})(typeof self!=='undefined'?self:globalThis);
