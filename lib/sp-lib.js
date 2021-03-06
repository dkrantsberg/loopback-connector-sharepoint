'use strict';
const _ = require('lodash');
const xml2js = require('xml2js');
const xmlBuilder = new xml2js.Builder({headless: true, renderOpts: {pretty: false}});

class SPLib {
  constructor(model) {
    this.model = model;
  }

  /**
   * Builds CAML corresponding to LoopBack filter object.
   * See documentation here: https://loopback.io/doc/en/lb3/Querying-data.html
   * @param filter LoopBack filter object
   * @returns {{ViewXml: string}}
   */
  buildQuery(filter) {
    return {
      ViewXml: `<View>${this.buildViewFields(filter.fields)}<Query>${this.buildWhere(filter.where)}${this.buildOrderBy(filter.order)}</Query>${this.buildRowLimit(filter.limit)}</View>`
    };
  }

  /**
   * Returns CAML XML query corresponding to LoopBack 'where' filter
   * @param lbWhere Loopback `where` filter object
   * @returns string CAML XML string
   */
  buildWhere(lbWhere) {
    if (_.isEmpty(lbWhere)) {
      return '';
    }
    const where = this._buildCamlCondition(lbWhere);
    return xmlBuilder.buildObject({Where: where});
  }

  /**
   * Builds CAML <OrderBy> tag from LoopBack order filter
   * @param order LoopBack order filter property. See more: https://loopback.io/doc/en/lb3/Order-filter.html
   * @returns {string} CAML XML string containing <OrderBy> tag
   */
  buildOrderBy(order) {
    // if order is not specified then order by ID descending
    if (_.isEmpty(order)) {
      return '<OrderBy><FieldRef Name="ID" Ascending="False"/></OrderBy>';
    }
    const fieldRefs = _.isArray(order) ?
      _.map(order, this._getOrderByFieldRef.bind(this)) : this._getOrderByFieldRef(order);
    const jsOrderBy = {
      OrderBy: fieldRefs
    };
    return xmlBuilder.buildObject(jsOrderBy);
  }

  /**
   * Builds CAML <ViewFields>
   * @param fields Properties of the LoopBack model to include
   * @returns {string}
   */
  buildViewFields(fields) {
    if (_.isEmpty(fields)) {
      return '';
    }
    const viewFields = _.map(fields, field => {
      return {FieldRef: {$: {Name: this.getSPFieldName(field)}}};
    });
    return xmlBuilder.buildObject({ViewFields: viewFields});
  }

  /**
   * Builds CAML condition corresponding to LoopBack 'where' filter
   * See documentation here: https://loopback.io/doc/en/lb3/Where-filter.html
   * @param lbWhere Loopback `where` filter object
   * @returns CAML string corresponding the passed `where` filter
   * @private
   */
  _buildCamlCondition(lbWhere) {
    const keys = Object.keys(lbWhere);
    if (keys.length === 0) {
      return {};
    } else if (keys.length === 1) {
      if ((keys[0] === 'or' || keys[0] === 'and') && _.isArray(lbWhere[keys[0]])) {
        // For complex conditions involving logical AND / OR operators
        // build CAML using recursion
        const camlOperator = getCamlName(keys[0]);
        const result = {[camlOperator]: {}};
        this._buildLogicalCaml(camlOperator, lbWhere[keys[0]], result);
        return result;
      } else {
        // for simple condition (w/o logical operators) just return CAML expression
        return this._buildCamlExpression(lbWhere);
      }
    } else {
      throw new Error('Invalid \'where\' clause. It must be in {key: value} format.');
    }
  }

  /**
   * Builds CAML filter condition corresponding to LoopBack compound logical filters.
   * For example:
   *      [{title: 'My Post'}, {content: 'Hello'}]
   *                  or
   *      [{and: [{ field1: 'foo' }, { field2: 'bar' }] }, { field1: 'morefoo' }]
   *
   * @param operator Logical operator, either And or Or
   * @param lbConditions An array of conditions
   * @param result A variable for accumulating the result
   * @param index Current position in the lbConditions. Used only for recursive calls.
   * @private
   */
  _buildLogicalCaml(operator, lbConditions, result, index) {
    index = index || 0;
    if (index === lbConditions.length) {
      return;
    }
    if (index === 0) {
      const firstCondition = this._buildCamlCondition(lbConditions[0]);
      const secondCondition = this._buildCamlCondition(lbConditions[1]);
      this._addCamlCondition(operator, result, firstCondition);
      this._addCamlCondition(operator, result, secondCondition);
      // call this method recursively starting with 3rd position
      this._buildLogicalCaml(operator, lbConditions, result, 2);
    } else {
      // CAML logical filers can contain only 2 conditions.
      // For 3 and more conditions they must be nested within each other
      const condition = this._buildCamlCondition(lbConditions[index]);
      const newCamlElement = this._addCamlCondition(operator, result, condition);
      this._buildLogicalCaml(operator, lbConditions, newCamlElement, index + 1);
    }
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
  }

  /**
   *  Builds CAML <RowLimit> tag
   * @param limit
   * @returns {string} CAML XML <RowLimit> tag
   */
  buildRowLimit(limit) {
    limit = _.parseInt(limit);
    if (!limit) {
      return '';
    }
    return xmlBuilder.buildObject({RowLimit: limit});
  }

  _buildCamlExpression(expression) {
    const {field, operator, value} = parseExpression(expression);
    const fieldType = this.getSPFieldType(field);
    if (operator === 'inq' || operator === 'nin') {
      if (!_.isArray(value)) {
        throw new Error('Invalid \'in\' values. Must be an array.');
      }
      const camlValues = _.map(value, (v) => ({_: v, $: {Type: fieldType}}));
      return {
        [getCamlName(operator)]: {
          FieldRef: {$: {Name: this.getSPFieldName(field)}},
          Values: {Value: camlValues}
        }
      };
    }
    let formattedValue;
    if (value instanceof Date) {
      formattedValue = value.toISOString();
    } else if (typeof value === 'boolean') {
      formattedValue = +value;
    } else {
      formattedValue = value;
    }
    return {
      [getCamlName(operator)]: {
        FieldRef: {$: {Name: this.getSPFieldName(field)}},
        Value: {
          _: formattedValue,
          $: {Type: fieldType}
        }
      }
    };
  }

  _addCamlCondition(newLogicalOperator, camlObject, camlCondition) {
    const logicalOperator = _.last(Object.keys(camlObject));
    const operator = Object.keys(camlCondition)[0];
    const condition = Object.values(camlCondition)[0];
    if (!camlObject[logicalOperator].hasOwnProperty(operator)) {
      // if this operator is not yet defined it then add it as single object
      camlObject[logicalOperator][operator] = condition;
    } else if (!Array.isArray(camlObject[logicalOperator][operator])) {
      // if this operator already present then turn it into the array and add new condition to it
      camlObject[logicalOperator][operator] = [camlObject[logicalOperator][operator], condition];
    } else if (Array.isArray(camlObject[logicalOperator][operator]) && camlObject[logicalOperator][operator].length === 2) {
      const lastCondition = camlObject[logicalOperator][operator].pop();
      camlObject[logicalOperator][operator] = camlObject[logicalOperator][operator][0];
      camlObject[logicalOperator][newLogicalOperator] = {[operator]: [lastCondition, condition]};
      return camlObject[logicalOperator];
    }
    // if resulting CAML object has 3 conditions then move the last 2 of them inside And/Or element
    if (Object.keys(camlObject[logicalOperator]).length === 3) {
      const keys = Object.keys(camlObject[logicalOperator]);
      const values = Object.values(camlObject[logicalOperator]);
      camlObject[logicalOperator] = {
        [keys[0]]: values[0],
        [logicalOperator]: {[keys[1]]: values[1], [keys[2]]: values[2]}
      };
      return camlObject[logicalOperator];
    }
  }

  getSPFieldName(property) {
    return _.get(this.model, `properties.${property}.sharepoint.columnName`) || property;
  }

  getSPFieldType(property) {
    if (property === 'ID') {
      return 'Number';
    }
    const propDefinition = this.model.properties[property];
    if (!propDefinition) {
      throw new Error(`Property ${property} is not defined for type ${this.model.name}.`);
    }
    return _.get(propDefinition, 'sharepoint.dataType') ||
      getDefaultSharePointType(_.get(propDefinition, 'type.name'));
  }
}

exports.SPLib = SPLib;

/*
* Parses loopback expression objects of types {field: value} or {field: {operator: value}}
* and returns a triple object with 3 distinct variables: field, operator, value
*/
function parseExpression(expression) {
  let operator;
  let value;

  const expressionKeys = Object.keys(expression);
  if (expressionKeys.length !== 1) {
    throw new Error(`Invalid expression: ${JSON.stringify(expression)}.`);
  }
  const field = expressionKeys[0];
  if (!_.isObject(expression[field])) {
    operator = '=';
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

function getCamlName(lbName) {
  switch (_.toLower(lbName)) {
    case 'and':
      return 'And';
    case 'or':
      return 'Or';
    case '=':
      return 'Eq';
    case 'neq':
      return 'Neq';
    case 'gt':
      return 'Gt';
    case 'gte':
      return 'Geq';
    case 'lt':
      return 'Lt';
    case 'lte':
      return 'Leq';
    case 'inq':
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
      return 'Number';
    case 'boolean':
      return 'Boolean';
    case 'date':
      return 'DateTime';
    default:
      return 'Text';
  }
}
