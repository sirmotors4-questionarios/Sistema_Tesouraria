/* Contrato de integração Tesouraria ↔ Excel/VBA. Não concede acesso ao SharePoint. */
(function(root,factory){const api=factory();if(typeof module==="object"&&module.exports)module.exports=api;root.TES_DATA=api})(typeof globalThis!=="undefined"?globalThis:this,function(){"use strict";
const CAMPOS=Object.freeze({
 requisicao:"Nr de Requisição",data:"Data",empresa:"Empresa",sector:"Sector",departamento:"Departamento",
 fornecedor:"Fornecedor",descricao:"Descrição",valor:"Valor",prioridade:"Prioridade",estado:"Status",
 categoria:"Categoria",tipoCusto:"Tipo de Custo",regularidade:"Regularidade",requisitadoPor:"Requisitado Por",
 observacoes:"Obs"
});
const CAMPOS_DESPACHO=Object.freeze({numero:"Nr de Despacho",data:"Data de Despacho",estado:"Estado do Despacho",dataSaida:"Data de Saída"});
const CAMPOS_PAGAMENTO=Object.freeze({forma:"F. de Pagamento",referencia:"Ref. Pagamento",banco:"Nome do Banco",conta:"Nr de Conta"});
const CAMPOS_DOCUMENTO=Object.freeze({requisicao:"R_Pagamento",suporte:"D_Suporte",comprovativo:"C_pagamento",recibo:"Recibo"});
const ESTADOS=Object.freeze(["Pendente","Aguardando Saída","Pago"]);
function string(v){return v==null?"":String(v).trim()}
function decimal(v){if(typeof v==="number")return Number.isFinite(v)?v:NaN;let s=string(v).replace(/\s/g,"");if(s.includes(",")&&s.includes("."))s=s.lastIndexOf(",")>s.lastIndexOf(".")?s.replace(/\./g,"").replace(",","."):s.replace(/,/g,"");else if(s.includes(","))s=s.replace(",",".");return /^-?\d+(\.\d+)?$/.test(s)?Number(s):NaN}
function isoDate(v){
 if(v==null||v==="")return "";
 const t=string(v);
 const n=typeof v==="number"?v:/^\d{4,6}(?:\.\d+)?$/.test(t)?Number(t):NaN;
 if(Number.isFinite(n)&&n>=1&&n<2958466)return new Date(Date.UTC(1899,11,30)+Math.floor(n)*86400000).toISOString().slice(0,10);
 let m=t.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T ]|$)/);
 if(m)return m[1]+"-"+m[2]+"-"+m[3];
 m=t.match(/^(\d{2})[\/.\-](\d{2})[\/.\-](\d{4})(?:\s|$)/);
 if(m){const d=new Date(Date.UTC(+m[3],+m[2]-1,+m[1]));return d.getUTCFullYear()===+m[3]&&d.getUTCMonth()===+m[2]-1&&d.getUTCDate()===+m[1]?m[3]+"-"+m[2]+"-"+m[1]:"";}
 return "";
}
function normalizeEstado(v){const s=string(v).toLocaleLowerCase("pt");const match=ESTADOS.find(x=>x.toLocaleLowerCase("pt")===s);return match||string(v)}
function fromExcel(row){if(!row||typeof row!=="object"||Array.isArray(row))throw Error("Registo Excel inválido");const out={};for(const [web,excel] of Object.entries(CAMPOS))out[web]=row[excel]??"";out.valor=decimal(out.valor);out.data=isoDate(out.data);out.dataOriginal=row.Data??null;out.estado=normalizeEstado(out.estado);out.numero=row.Nr??"";for(const [k,col] of Object.entries(CAMPOS_DESPACHO))out["despacho_"+k]=row[col]??"";for(const [k,col] of Object.entries(CAMPOS_PAGAMENTO))out["pagamento_"+k]=row[col]??"";for(const [k,col] of Object.entries(CAMPOS_DOCUMENTO))out["documento_"+k]=row[col]??"";out.id=string(row.ID_Despesa||"");return out}
function validateWeb(row){const errors=[];for(const k of ["requisicao","data","empresa","fornecedor","descricao"])if(!string(row[k]))errors.push(k+" obrigatório");if(!Number.isFinite(decimal(row.valor))||decimal(row.valor)<=0)errors.push("valor inválido");if(row.estado&&!ESTADOS.includes(row.estado))errors.push("estado não reconhecido (não normalizar silenciosamente)");if(!/^\d{4}-\d{2}-\d{2}$/.test(string(row.data)))errors.push("data inválida");return errors}
function toExcel(row){const errors=validateWeb(row);if(errors.length)throw Error(errors.join("; "));const out={};for(const [web,excel] of Object.entries(CAMPOS)){if(Object.prototype.hasOwnProperty.call(row,web))out[excel]=web==="valor"?decimal(row.valor):row[web]}return out}
function contractCheck(headers){const set=new Set(headers);return {missing:Object.values(CAMPOS).filter(x=>!set.has(x)),available:Object.values(CAMPOS).filter(x=>set.has(x)),unknown:headers.filter(x=>!Object.values(CAMPOS).includes(x))}}
function unwrapConsulta(response){const result=response?.registos??response?.body?.registos??response?.value;if(!Array.isArray(result))throw Error("A resposta da consulta não contém uma lista de registos");return result.map(fromExcel)}
return Object.freeze({CAMPOS,CAMPOS_DESPACHO,CAMPOS_PAGAMENTO,CAMPOS_DOCUMENTO,ESTADOS,fromExcel,toExcel,contractCheck,unwrapConsulta,validateWeb,isoDate,decimal});
});
