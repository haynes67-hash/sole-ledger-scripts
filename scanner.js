/**
 * ═══════════════════════════════════════════════════════════════
 *  SOLE LEDGER — Gmail Sales Tracker (Google Apps Script)
 * ═══════════════════════════════════════════════════════════════
 *
 *  SETUP INSTRUCTIONS
 *  ──────────────────
 *  1. Go to https://script.google.com and click "New project"
 *  2. Paste this entire file into the editor, replacing any existing code
 *  3. Click "Save" (Ctrl+S / Cmd+S)
 *  4. Click "Run" → choose "setup" → accept permissions
 *  5. Click "Run" → choose "installTrigger"
 *  6. Deploy as Web App:
 *     • Click "Deploy" → "New deployment"
 *     • Type: Web app
 *     • Execute as: Me
 *     • Who has access: Anyone
 *     • Click "Deploy" → copy the Web App URL
 *  7. Paste the Web App URL into Sole Ledger → Sales → Gmail Sync
 * ═══════════════════════════════════════════════════════════════
 */

var SHEET_NAME  = "Sales";
var SPREAD_NAME = "Sole Ledger Sales";
var DAYS_BACK   = 90;
var COLS = ["id","date","platform","brand","model","colorway","sku","size","salePrice","cost","notes","emailSubject","emailDate","raw"];

// ── Web App endpoint — called by Sole Ledger app ───────────────
function doGet(e) {
  try {
    // Run a fresh scan first
    scanEmails();
    // Return all sales as JSON
    var ss    = getOrCreateSheet();
    var sheet = ss.getSheetByName(SHEET_NAME);
    var last  = sheet.getLastRow();
    var sales = [];
    if (last > 1) {
      var data = sheet.getRange(2, 1, last - 1, COLS.length).getValues();
      data.forEach(function(row) {
        var obj = {};
        COLS.forEach(function(c, i) { obj[c] = row[i] || ""; });
        if (obj.id) sales.push(obj);
      });
    }
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, sales: sales, count: sales.length }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Entry point (runs every 15 min via trigger) ────────────────
function scanEmails() {
  var ss    = getOrCreateSheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  var existing = getExistingIds(sheet);
  var cutoff   = getCutoffDate(sheet);

  var results = [];
  results = results.concat(scanStockX(cutoff, existing));
  results = results.concat(scanGOAT(cutoff, existing));
  results = results.concat(scanEbay(cutoff, existing));
  results = results.concat(scanDepop(cutoff, existing));
  results = results.concat(scanGeneric(cutoff, existing));

  if (results.length > 0) {
    results.forEach(function(row) {
      sheet.appendRow(COLS.map(function(c){ return row[c] || ""; }));
    });
    Logger.log("Added " + results.length + " new sales.");
  } else {
    Logger.log("No new sales found.");
  }

  var meta = ss.getSheetByName("_meta") || ss.insertSheet("_meta");
  meta.getRange("A1").setValue("Last sync");
  meta.getRange("B1").setValue(new Date().toISOString());
  meta.getRange("A2").setValue("Total sales");
  meta.getRange("B2").setValue(sheet.getLastRow() - 1);
}

// ── StockX ─────────────────────────────────────────────────────
function scanStockX(cutoff, existing) {
  var results = [];
  var threads = GmailApp.search(
    'from:(noreply@stockx.com OR orders@stockx.com) subject:("Your sale" OR "sold" OR "Sale confirmed") after:' + dateStr(cutoff),
    0, 50
  );
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getDate() < cutoff) return;
      var id  = "stockx_" + msg.getId();
      if (existing[id]) return;
      var body = msg.getPlainBody();
      var subj = msg.getSubject();
      var price = extractPrice(body, [
        /sale\s+price[\s\S]{0,40}?\$\s*([\d,]+\.?\d*)/i,
        /you\s+sold\s+for\s+\$\s*([\d,]+\.?\d*)/i,
        /payout[\s\S]{0,30}?\$\s*([\d,]+\.?\d*)/i,
      ]);
      var shoe = extractShoe(body, subj, [
        /congratulations[^!]*!\s*([\w].*?)\s*(?:in size|\n|has sold)/i,
        /you\s+sold\s+(?:a\s+)?([\w].*?)\s+(?:in size|\()/i,
        /item[:\s]+([\w].*?)\n/i
      ]);
      var size = extractSize(body, [/size\s*:?\s*([0-9]+\.?[0-9]*)/i, /in\s+size\s+([0-9]+\.?[0-9]*)/i]);
      if (!price && !shoe) return;
      var parsed = parseShoe(shoe);
      results.push({ id:id, date:fmtDate(msg.getDate()), platform:"StockX", brand:parsed.brand, model:parsed.model, colorway:parsed.colorway, sku:"", size:size, salePrice:price, cost:"", notes:"", emailSubject:subj, emailDate:msg.getDate().toISOString(), raw:body.substring(0,500) });
    });
  });
  return results;
}

// ── GOAT ───────────────────────────────────────────────────────
function scanGOAT(cutoff, existing) {
  var results = [];
  var threads = GmailApp.search(
    'from:(orders@goat.com OR noreply@goat.com) subject:("sold" OR "sale" OR "order confirmed") after:' + dateStr(cutoff),
    0, 50
  );
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getDate() < cutoff) return;
      var id = "goat_" + msg.getId();
      if (existing[id]) return;
      var body = msg.getPlainBody();
      var subj = msg.getSubject();
      var price = extractPrice(body, [/total[\s\S]{0,20}?\$\s*([\d,]+\.?\d*)/i, /payout[\s\S]{0,20}?\$\s*([\d,]+\.?\d*)/i, /\$\s*([\d,]+\.?\d*)/i]);
      var shoe  = extractShoe(body, subj, [/item[:\s]+([\w].*?)\n/i, /product[:\s]+([\w].*?)\n/i]);
      var size  = extractSize(body, [/size\s*:?\s*([0-9]+\.?[0-9]*)/i]);
      if (!price && !shoe) return;
      var parsed = parseShoe(shoe);
      results.push({ id:id, date:fmtDate(msg.getDate()), platform:"GOAT", brand:parsed.brand, model:parsed.model, colorway:parsed.colorway, sku:"", size:size, salePrice:price, cost:"", notes:"", emailSubject:subj, emailDate:msg.getDate().toISOString(), raw:body.substring(0,500) });
    });
  });
  return results;
}

// ── eBay ───────────────────────────────────────────────────────
function scanEbay(cutoff, existing) {
  var results = [];
  var threads = GmailApp.search(
    'from:(ebay.com) subject:("You sold" OR "sold on eBay" OR "item sold") after:' + dateStr(cutoff),
    0, 50
  );
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getDate() < cutoff) return;
      var id = "ebay_" + msg.getId();
      if (existing[id]) return;
      var body = msg.getPlainBody();
      var subj = msg.getSubject();
      var price = extractPrice(body, [/sale\s+price[\s\S]{0,20}?\$\s*([\d,]+\.?\d*)/i, /sold\s+for\s+\$\s*([\d,]+\.?\d*)/i, /total[\s\S]{0,20}?\$\s*([\d,]+\.?\d*)/i]);
      var shoe = subj.replace(/^you sold[:\s]*/i,"").replace(/\s*-\s*ebay.*$/i,"").trim();
      if (!shoe) shoe = extractShoe(body, subj, [/item[:\s]+([\w].*?)\n/i]);
      var size = extractSize(body, [/size\s*:?\s*([0-9]+\.?[0-9]*)/i]);
      if (!price) return;
      var parsed = parseShoe(shoe);
      results.push({ id:id, date:fmtDate(msg.getDate()), platform:"eBay", brand:parsed.brand, model:parsed.model, colorway:parsed.colorway, sku:"", size:size, salePrice:price, cost:"", notes:"", emailSubject:subj, emailDate:msg.getDate().toISOString(), raw:body.substring(0,500) });
    });
  });
  return results;
}

// ── Depop ──────────────────────────────────────────────────────
function scanDepop(cutoff, existing) {
  var results = [];
  var threads = GmailApp.search(
    'from:(no-reply@depop.com) subject:("sold" OR "You made a sale") after:' + dateStr(cutoff),
    0, 50
  );
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getDate() < cutoff) return;
      var id = "depop_" + msg.getId();
      if (existing[id]) return;
      var body = msg.getPlainBody();
      var subj = msg.getSubject();
      var price = extractPrice(body, [/\$\s*([\d,]+\.?\d*)/i, /£\s*([\d,]+\.?\d*)/i]);
      var shoe  = extractShoe(body, subj, [/item[:\s]+([\w].*?)\n/i]);
      var size  = extractSize(body, [/size\s*:?\s*([0-9]+\.?[0-9]*)/i]);
      if (!price) return;
      var parsed = parseShoe(shoe);
      results.push({ id:id, date:fmtDate(msg.getDate()), platform:"Depop", brand:parsed.brand, model:parsed.model, colorway:parsed.colorway, sku:"", size:size, salePrice:price, cost:"", notes:"", emailSubject:subj, emailDate:msg.getDate().toISOString(), raw:body.substring(0,500) });
    });
  });
  return results;
}

// ── Generic ────────────────────────────────────────────────────
function scanGeneric(cutoff, existing) {
  var results = [];
  var threads = GmailApp.search(
    'subject:("your sneaker sold" OR "shoe sale confirmed" OR "order shipped to buyer") after:' + dateStr(cutoff),
    0, 30
  );
  threads.forEach(function(thread) {
    thread.getMessages().forEach(function(msg) {
      if (msg.getDate() < cutoff) return;
      var id = "generic_" + msg.getId();
      if (existing[id]) return;
      var body = msg.getPlainBody();
      var subj = msg.getSubject();
      var price = extractPrice(body, [/\$\s*([\d,]+\.?\d*)/i]);
      if (!price) return;
      var shoe = extractShoe(body, subj, []);
      var size = extractSize(body, [/size\s*:?\s*([0-9]+\.?[0-9]*)/i]);
      var parsed = parseShoe(shoe);
      results.push({ id:id, date:fmtDate(msg.getDate()), platform:guessFromDomain(msg.getFrom()), brand:parsed.brand, model:parsed.model, colorway:parsed.colorway, sku:"", size:size, salePrice:price, cost:"", notes:"", emailSubject:subj, emailDate:msg.getDate().toISOString(), raw:body.substring(0,500) });
    });
  });
  return results;
}

// ── Helpers ────────────────────────────────────────────────────
function extractPrice(body, patterns) {
  for (var i=0; i<patterns.length; i++) { var m=body.match(patterns[i]); if(m) return parseFloat(m[1].replace(/,/g,"")); }
  return null;
}
function extractShoe(body, subj, patterns) {
  for (var i=0; i<patterns.length; i++) { var m=body.match(patterns[i]); if(m&&m[1]&&m[1].trim().length>2) return m[1].trim(); }
  return subj||"";
}
function extractSize(body, patterns) {
  for (var i=0; i<patterns.length; i++) { var m=body.match(patterns[i]); if(m) return parseFloat(m[1]); }
  return "";
}
var KNOWN_BRANDS = ["Nike","Adidas","Jordan","New Balance","Yeezy","Puma","Reebok","Converse","Vans","Asics","Salomon","On Running","Hoka","Brooks","Saucony","Under Armour","Balenciaga","Gucci","Louis Vuitton","Dior","Off-White","Travis Scott"];
function parseShoe(name) {
  if (!name) return {brand:"",model:"",colorway:""};
  name=name.trim();
  for (var i=0; i<KNOWN_BRANDS.length; i++) {
    if (name.toLowerCase().startsWith(KNOWN_BRANDS[i].toLowerCase())) {
      var rest=name.slice(KNOWN_BRANDS[i].length).trim();
      var colorway=""; var cparen=rest.match(/\(([^)]+)\)$/);
      if (cparen){colorway=cparen[1];rest=rest.replace(cparen[0],"").trim();}
      return {brand:KNOWN_BRANDS[i],model:rest,colorway:colorway};
    }
  }
  var parts=name.split(" "); return {brand:parts[0]||"",model:parts.slice(1).join(" ")||"",colorway:""};
}
function guessFromDomain(from) {
  if(/stockx/i.test(from)) return "StockX"; if(/goat/i.test(from)) return "GOAT";
  if(/ebay/i.test(from)) return "eBay"; if(/depop/i.test(from)) return "Depop";
  if(/klekt/i.test(from)) return "Klekt"; return "Other";
}
function fmtDate(d) { return Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy-MM-dd"); }
function dateStr(d) { return Utilities.formatDate(d,Session.getScriptTimeZone(),"yyyy/MM/dd"); }
function getCutoffDate(sheet) {
  var last=sheet.getLastRow();
  if (last>1) {
    var dateCol=COLS.indexOf("emailDate")+1;
    var dates=sheet.getRange(2,dateCol,last-1,1).getValues().flat().filter(Boolean);
    if (dates.length>0) { return dates.map(function(d){return new Date(d);}).sort(function(a,b){return b-a;})[0]; }
  }
  var d=new Date(); d.setDate(d.getDate()-DAYS_BACK); return d;
}
function getExistingIds(sheet) {
  var map={}; var last=sheet.getLastRow(); if(last<2) return map;
  var idCol=COLS.indexOf("id")+1;
  sheet.getRange(2,idCol,last-1,1).getValues().forEach(function(r){if(r[0])map[r[0]]=true;});
  return map;
}
function getOrCreateSheet() {
  var files=DriveApp.getFilesByName(SPREAD_NAME);
  if(files.hasNext()) return SpreadsheetApp.open(files.next());
  var ss=SpreadsheetApp.create(SPREAD_NAME);
  var sheet=ss.getActiveSheet().setName(SHEET_NAME);
  sheet.appendRow(COLS); sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,COLS.length).setFontWeight("bold").setBackground("#1e2130").setFontColor("white");
  Logger.log("Created sheet: "+ss.getUrl()); return ss;
}
function setup() {
  var ss=getOrCreateSheet();
  Logger.log("Sheet URL: "+ss.getUrl());
  Logger.log("Sheet ID: "+ss.getId());
  scanEmails();
}
function installTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger("scanEmails").timeBased().everyMinutes(15).create();
  Logger.log("Trigger installed.");
}
function rescan() {
  var ss=getOrCreateSheet(); var sheet=ss.getSheetByName(SHEET_NAME);
  if(sheet.getLastRow()>1) sheet.getRange(2,1,sheet.getLastRow()-1,COLS.length).clearContent();
  var d=new Date(); d.setDate(d.getDate()-DAYS_BACK);
  var results=[].concat(scanStockX(d,{}),scanGOAT(d,{}),scanEbay(d,{}),scanDepop(d,{}));
  results.forEach(function(row){sheet.appendRow(COLS.map(function(c){return row[c]||"";}));});
  Logger.log("Rescan complete. Found "+results.length+" sales.");
}
