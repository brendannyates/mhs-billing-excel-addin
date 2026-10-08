import test from 'node:test';import assert from 'node:assert/strict';
import {scopeTickets} from '../core.js';
const t={Id:31,'Completion time':'2026-10-02','Patient Clinic:':'San Rafael - CA','MRN:':'00123','Patient:':'TEST','Status:':'MHS - Submitted to Arietis',Email:'owner@example.com','Submitted to Vendor At':'2026-10-02'};
test('scope defaults to assigned clinics, My Tickets stays within clinic scope',()=>{const other={...t,Id:32,'Patient Clinic:':'Castro - CA'};const p={email:'owner@example.com',clinics:['San Rafael - CA']};assert.equal(scopeTickets([t,other],p).length,1);assert.equal(scopeTickets([t,other],p,{mine:true}).length,1);assert.equal(scopeTickets([t,other],p,{scope:'all',mrn:'00123'}).length,2);});
test('no assigned clinics shows every clinic instead of an empty dashboard',()=>{const other={...t,Id:32,'Patient Clinic:':'Castro - CA'};assert.equal(scopeTickets([t,other],{email:'x@example.com',clinics:[]}).length,2);});
test('My Tickets means created by me, even after owner is reassigned',()=>{const p={email:'owner@example.com',clinics:['San Rafael - CA']};assert.equal(scopeTickets([{...t,'Owner Email':'someone.else@example.com'}],p,{mine:true}).length,1);assert.equal(scopeTickets([{...t,Email:'other@example.com','Owner Email':'owner@example.com'}],p,{mine:true}).length,0);});

