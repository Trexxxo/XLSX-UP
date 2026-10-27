"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const test = require("node:test");

const source = fs.readFileSync(path.resolve(__dirname, "../app.js"), "utf8");
const mockElement = () => ({
  addEventListener() {},
  classList: { add() {}, remove() {} },
  style: {},
  replaceChildren() {},
  setAttribute() {},
  removeAttribute() {},
  append() {}
});
const elements = new Map();
const context = {
  document: {
    getElementById(id) {
      if (!elements.has(id)) elements.set(id, mockElement());
      return elements.get(id);
    }
  },
  window: { addEventListener() {} },
  console
};
vm.createContext(context);
vm.runInContext(source, context);
const transform = vm.runInContext("stripSheetProtection", context);

test("removes a regular self-closing protection element only", () => {
  const original = '<?xml version="1.0"?><worksheet><sheetData><c r="A1"><v>25</v></c></sheetData><sheetProtection sheet="1" password="DABC"/><mergeCells count="1"/></worksheet>';
  const result = transform(original);
  assert.equal(result.removed, 1);
  assert.equal(result.xml, original.replace('<sheetProtection sheet="1" password="DABC"/>', ""));
});

test("removes all worksheet protection instances without changing other XML", () => {
  const original = '<worksheet><sheetProtection algorithmName="SHA-512" hashValue="some+base64" /><sheetData/><sheetProtection sheet="1" /></worksheet>';
  const result = transform(original);
  assert.equal(result.removed, 2);
  assert.equal(result.xml, "<worksheet><sheetData/></worksheet>");
});

test("handles > inside quoted tag attributes", () => {
  const result = transform('<worksheet><sheetProtection password="x>y" /><sheetData/></worksheet>');
  assert.equal(result.removed, 1);
  assert.equal(result.xml, "<worksheet><sheetData/></worksheet>");
});

test("supports a namespace prefix and paired protection tags", () => {
  const result = transform('<worksheet><x:sheetProtection x:sheet="1"></x:sheetProtection><sheetData/></worksheet>');
  assert.equal(result.removed, 1);
  assert.equal(result.xml, "<worksheet><sheetData/></worksheet>");
});

test("does not remove similar tag names", () => {
  const original = '<worksheet><sheetProtectionExtra/><notSheetProtection/><sheetData/></worksheet>';
  const result = transform(original);
  assert.equal(result.removed, 0);
  assert.equal(result.xml, original);
});

test("passes through an unprotected worksheet", () => {
  const original = '<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>';
  assert.equal(transform(original).xml, original);
  assert.equal(transform(original).removed, 0);
});

test("rejects malformed protection tags rather than damaging a sheet", () => {
  assert.throws(() => transform('<worksheet><sheetProtection sheet="1"'), /incomplete XML tag/);
  assert.throws(() => transform('<worksheet><sheetProtection><sheetData/></worksheet>'), /Malformed sheetProtection/);
});
