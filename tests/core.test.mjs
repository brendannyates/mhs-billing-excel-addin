import test from 'node:test';import assert from 'node:assert/strict';
import {merge} from '../core.js';
test('Forms blank overwrite preserves ticket 31 fields',()=>{const m={'Id':31,'Patient Clinic:':'San Rafael - CA','MRN:':'00123','Notes':'Keep this'};assert.deepEqual(merge(m,{Id:31,'Patient Clinic:':'','MRN:':'',Notes:''},m).next,m);});
test('unchanged raw does not reverse Master edits',()=>{const b={Id:1,Notes:'Old'};assert.equal(merge({...b,Notes:'Staff edit'},b,b).next.Notes,'Staff edit');});
test('changed filled source updates unchanged Master',()=>{assert.equal(merge({Id:1,Notes:'Old'},{Id:1,Notes:'New'},{Id:1,Notes:'Old'}).next.Notes,'New');});
test('simultaneous field edit retains Master and flags conflict',()=>{const x=merge({Notes:'Staff'},{Notes:'Forms'},{Notes:'Old'});assert.equal(x.next.Notes,'Staff');assert.deepEqual(x.conflicts,['Notes']);});
test('first sync differing populated fields need review',()=>{assert.deepEqual(merge({Notes:'Master'},{Notes:'Raw'}).conflicts,['Notes']);});
test('zero is a filled value',()=>assert.equal(merge({Amount:2},{Amount:0},{Amount:2}).next.Amount,0));

