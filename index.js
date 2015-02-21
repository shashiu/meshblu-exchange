'use strict';
var util = require('util');
var EventEmitter = require('events').EventEmitter;
var debug = require('debug')('meshblu-exchange');
var ews = require('ews');

var MESSAGE_SCHEMA = {
  type: 'object',
  properties: {
    command: {
      type: 'string',
      enum: ['GetCalendarEntries'], // add more here
      default: 'GetCalendarEntries',
      required: true
    },
    mailbox: {
      type: 'string',
      required: true
    },
    startDate: {
      type: 'string',
      required: true
    },
    endDate: {
      type: 'string',
      required: true
    }
  }
};

var OPTIONS_SCHEMA = {
};

function Plugin(){
  this.options = {};
  this.messageSchema = MESSAGE_SCHEMA;
  this.optionsSchema = OPTIONS_SCHEMA;
  return this;
}
util.inherits(Plugin, EventEmitter);

Plugin.prototype.onMessage = function(message){
  var payload = message.payload;
  var self = this;
  if (payload.command == 'GetCalendarEntries') {
    this._getCalendarEntries(payload.mailbox, payload.startDate, payload.endDate, function (entries) {
      self.emit('message', { devices: ['*'], topic: 'entries', payload: entries });
    });
  }
};

Plugin.prototype.onConfig = function(device){
  this.setOptions(device.options||{});
};

Plugin.prototype.setOptions = function(options){
  this.options = options;
  this.setupEws();
};

module.exports = {
  messageSchema: MESSAGE_SCHEMA,
  optionsSchema: OPTIONS_SCHEMA,
  Plugin: Plugin
};

Plugin.prototype.setupEws = function() {
  var self = this;

  /* For now we are getting exchange credentials from the process environment */
  self._session = new ews.MSExchange(
    process.env.EWSUSER,
    process.env.EWSPASS,
    process.env.EWSDOMAIN
  );

  self._session.autoDiscover(process.env.EWSMAIL).then(function () {
    console.log("Exchange authenticated and autodiscovered EWS successfully");
  });

  self._getCalendarEntries = function (mailbox, startDate, endDate, callback) {
    var calendar = new ews.Calendar(self._session, mailbox);
    calendar.getEntries(startDate, endDate).then(callback);
  }
}