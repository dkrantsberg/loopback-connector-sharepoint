'use strict';
const _ = require('lodash');
const xml2js = require('xml2js');
const xmlBuilder = new xml2js.Builder({headless: true, renderOpts: {pretty: false}});

class CamlBuilder {
  constructor(model) {
    this.model = model;
    this.camlDebug = [];
  }

  buildQuery(filter) {
    return {
      ViewXml: `<View>${this.buildFields(filter.fields)}<Query>${this.buildWhere(filter.where)}${this.buildOrderBy(filter.order)}</Query>${this.buildRowLimit(filter.limit)}</View>`,
    };
  };

  buildOrderBy(order) {
    if (_.isEmpty(order)) {
      return;
    }
    const fieldRefs = _.isArray(order) ? _.map(order, this._getOrderByFieldRef) : this._getOrderByFieldRef(order);
    const jsOrderBy = {
      OrderBy: fieldRefs
    }
  };

  buildFields(fields) {
    if (_.isEmpty(fields)) {
      return;
    }
    const keys = Object.keys(fields);
    const values = Object.values(fields);
    const operation = _.uniq(values);

    let fieldsToInclude;
    if (operation[0] && operation.length === 1) {
      fieldsToInclude = keys;
    } else if (!operation[0] && operation.length === 1) {
      const modelFields = Object.keys(this.model.properties);
      fieldsToInclude = _.xor(modelFields, keys);
    } else {
      throw new Error('Invalid fields expression. All specified fields must be either included (true) or excluded (false)');
    }
    const viewFields = _.map(fieldsToInclude, fieldName => {
      return {FieldRef: {$: {Name: this.getSPFieldName(fieldName)}}};
    });
    return {ViewFields: viewFields};
  }

  _getOrderByFieldRef(orderClause) {
    if (!_.isString(orderClause)) {
      throw new Error('Invalid order expression. Must be a string.');
    }
    const clauseParts = orderClause.split(' ');
    if (clauseParts.length === 1) {
      return {FieldRef: {$: {Name: this.getSPFieldName(clauseParts[0])}}};
    } else if (clauseParts.length === 2) {
      let isAscending;
      if (clauseParts[1].toUpperCase() === 'ASC') {
        isAscending = 'TRUE';
      } else if (clauseParts[1].toUpperCase() === 'DESC') {
        isAscending = 'FALSE';
      } else {
        throw new Error('Invalid order direction. Must be either ASC or DESC');
      }
      return {FieldRef: {$: {Name: this.getSPFieldName(clauseParts[0]), Ascending: isAscending}}};
    }
  };

  buildRowLimit(limit) {
    limit = _.parseInt(limit);
    if (!limit) {
      return;
    }
    return xmlBuilder.buildQuery({RowLimit: limit});
  }

  buildWhere(lbWhere) {
    if (_.isEmpty(lbWhere)) {
      return;
    }
    const _where = _.cloneDeep(lbWhere);
    const jsCaml = this._buildWhere(_where);
    return xmlBuilder.buildObject(jsCaml);
  };


  _buildWhere(lbClause, camlClause, wrapElement) {
    this.camlDebug.push(_.cloneDeep(camlClause));
    camlClause = camlClause || {};
    if (!_.isArray(lbClause)) {
      const keys = Object.keys(lbClause);
      if (keys.length === 0) {
        return {};
      } else if (keys.length === 1) {
        // if this is a single condition then return Caml
        const value = lbClause[keys[0]];
        if (keys[0] === 'or' || keys[0] === 'and' || _.isArray(value)) {
          const camlKey = getCamlName(keys[0]);
          this._buildWhere(value, camlClause, camlKey)
        } else {
          return this._buildCamlExpression(lbClause);
        }
      } else {
        // otherwise treat it as chain of 'and'
      }
    } else if (lbClause.length && wrapElement) {
      const expression = lbClause.shift();
      if (isLogical(expression)) {
        camlClause[wrapElement] = this._buildWhere(expression, _.cloneDeep(camlClause));
      } else {
        camlClause[wrapElement] = this._buildCamlExpression(expression);
      }
      this._buildWhere(lbClause, camlClause[wrapElement], wrapElement);
    }
    return camlClause;
  };

  _buildCamlExpression(expression) {
    const {field, operator, value} = parseExpression(expression);
    const fieldType = this.getSPFieldType(field);
    if (operator === 'in') {
      if (!_.isArray(value)) {
        throw new Error(`Invalid 'in' values. Must be an array.`);
      }
      const camlValues = _.map(value, (v) => ({Value: {_: v, $: {Type: fieldType}}}));
      return {
        [getCamlName(operator)]: {
          FieldRef: {$: {Name: this.getSPFieldName(field)}},
          Values: camlValues
        }
      };
    }
    return {
      [getCamlName(operator)]: {
        FieldRef: {$: {Name: field}},
        Value: {
          _: value,
          $: {Type: fieldType}
        }
      }
    };
  }

  getSPFieldName(property) {
    return _.get(this.model, `properties.${property}.sharepoint.columnName`) || property;
  }

  getSPFieldType(property) {
    return _.get(this.model, `properties.${property}.sharepoint.dataType`) ||
      getDefaultSharePointType(_.get(this.model, `properties.${property}.type.name`));
  }
}

exports.CamlBuilder = CamlBuilder;

/*
* Parses loopback expression objects of types {field: value} or {field: {operator: value}}
* and returns a triple object with 3 distinct variables: field, operator, value
*/
function parseExpression(expression) {
  let field;
  let operator;
  let value;

  const expressionKeys = Object.keys(expression);
  if (expressionKeys.length !== 1) {
    throw new Error(`Invalid expression: ${JSON.stringify(expression)}.`);
  }
  field = expressionKeys[0];
  if (!_.isObject(expression[field])) {
    operator = 'eq';
    value = expression[field];
  } else {
    const conditionKeys = Object.keys(expression[field]);
    if (conditionKeys.length !== 1) {
      throw new Error(`Invalid condition: ${JSON.stringify(expression[field])}.`);
    }
    operator = Object.keys(expression[field])[0];
    value = expression[field][operator];
  }
  return {field, operator, value};
}

function isLogical(expression) {
  if (_.isObject(expression)) {
    const keys = Object.keys(expression);
    if (keys.length === 1 &&
      (keys[0] === 'and' || keys[0] === 'or')
      && _.isArray(expression[keys[0]])) {
      return true;
    }
  }
  return false;
}

function getCamlName(lbName) {
  switch (_.toLower(lbName)) {
    case 'and':
      return 'And';
    case 'or':
      return 'Or';
    case 'eq':
      return 'Eq';
    case 'neq':
      return '';
    case 'Neq':
      return '';
    case 'gt':
      return '';
    case 'Gt':
      return '';
    case 'gte':
      return '';
    case 'Geq':
      return '';
    case 'lt':
      return '';
    case 'Lt':
      return '';
    case 'lte':
      return '';
    case 'Leq':
      return '';
    case 'inc':
      return '';
    case 'Includes':
      return '';
    case 'nin':
      return '';
    case 'NotIncludes':
      return '';
    case 'in':
      return 'In';
    case 'like':
      return 'BeginsWith';
    case 'contains':
      return 'Contains';
  }
}

function getDefaultSharePointType(lsType) {
  switch (lsType.toLowerCase()) {
    case 'string':
      return 'Text';
    case 'number':
      return 'Integer';
    case 'boolean':
      return 'Boolean';
    case 'date':
      return 'DateTime';
    default:
      return 'Text';
  }
}