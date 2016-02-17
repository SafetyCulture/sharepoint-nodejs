'use strict';

Object.defineProperty(exports, "__esModule", {
  value: true
});
exports.listType = undefined;
exports.toResponse = toResponse;
exports.convertItem = convertItem;
exports.toItem = toItem;
exports.toAudit = toAudit;

var _lodash = require('lodash');

var _lodash2 = _interopRequireDefault(_lodash);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var listType = exports.listType = function listType(name) {
  return 'SP.Data.' + name.replace(/ /g, '_x0020_') + 'Item';
};

/**
* Transforms a SafetyCulture 'responses' field to a response on SharePoint
* @param {object} responses Responses field from SafetyCulture
* @returns {object} Transformed SharePoint response
*/
function toResponse(responses) {
  var type = _lodash2.default.keys(responses)[0];

  switch (type) {
    case 'selected':
      var response = responses[type][0];
      return {
        'R Id': response.id,
        'R Type': response.type,
        'R Label': String(response.label),
        'R Short Label': response.short_label,
        'R Colour': response.colour,
        'R Image': response.image,
        'R Enable Score': response.enable_score,
        'R Score': response.score
      };
    case 'text':
      return {
        'R Type': 'text',
        'R Label': String(responses.text)
      };
    default:
      return {
        'R Type': type,
        'R Label': String(responses[type])
      };
  }
}

function convertItem(item, listName) {
  var type = listType(listName);
  var meta = { '__metadata': { 'type': type } };
  return _lodash2.default.merge({}, item, meta);
}

/**
* Transforms a SafetyCulture item to an item on SharePoint
* @param {object} scItem Item from SafetyCulture
* @returns {object} Transformed SharePoint item
*/
function toItem(scItem) {
  var item = {
    '__metadata': { 'type': 'SP.Data.SafetyCulture_x0020_ItemsListItem' },
    'Title': scItem.label,
    'Parent Id': scItem.parent_id,
    'Item Id': scItem.item_id,
    'Label': String(scItem.label),
    'Type': scItem.type
  };

  if (scItem.scoring) {
    item = _lodash2.default.assign(item, {
      'Score': scItem.scoring.score,
      'Max Score': scItem.scoring.max_score,
      'Percentage': scItem.scoring.score_percentage
    });

    if (scItem.scoring.combined) {
      item = _lodash2.default.assign(item, {
        'C Score': scItem.scoring.combined_score,
        'C Max Score': scItem.scoring.combined.max_score,
        'C Score Percentage': scItem.scoring.combined_score_percentage
      });
    }
  }

  if (scItem.responses) {
    item = _lodash2.default.assign(item, toResponse(scItem.responses));
  }

  if (scItem.options) {
    item = _lodash2.default.assign(item, {
      'O Weighting': scItem.options.weighting
    });
  }

  // remove all falsy values
  return _lodash2.default.pick(item, function (f) {
    return _lodash2.default.identity(f) || f === 0;
  });
}

/**
* Transforms a SafetyCulture audit to an audit on SharePoint
* @param {object} scAudit Audit from SafetyCulture
* @param {array} itemIds Array of Item Ids from SharePoint that relate to this audit
* @returns {object} Transformed SharePoint Audit
*/
function toAudit(scAudit, itemIds) {
  var auditTitle = function auditTitle(audit) {
    return audit.audit_data.name ? audit.audit_data.name : 'Audit ' + audit.template_data.metadata.name;
  };

  return _lodash2.default.pick({
    '__metadata': { 'type': 'SP.Data.SafetyCulture_x0020_AuditsListItem' },
    'Audit Id': scAudit.audit_id,
    'Title': auditTitle(scAudit),
    'Score': scAudit.audit_data.score,
    'Total Score': scAudit.audit_data.total_score,
    'Score Percentage': scAudit.audit_data.score_percentage,
    'Duration': scAudit.audit_data.duration,
    'Date Modified': scAudit.audit_data.date_modified,
    'Date Started': scAudit.audit_data.date_started,
    'Date Completed': scAudit.audit_data.date_completed,
    'SafetyCulture Owner': scAudit.audit_data.authorship.owner,
    'SafetyCulture Author': scAudit.audit_data.authorship.author,
    'Device Id': scAudit.audit_data.authorship.device_id,
    'Template Name': scAudit.template_data.metadata.name,
    'Template Description': scAudit.template_data.metadata.description,
    'ItemsId': { results: itemIds }
  }, function (f) {
    return _lodash2.default.identity(f) || f === 0;
  });
}