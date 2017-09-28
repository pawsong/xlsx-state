import { WorkBook } from 'xlsx';
import cloneDeep = require('lodash/cloneDeep');

// +---------------------+
// | FORMULAS REGISTERED |
// +---------------------+
const xlsx_Fx = {
  'FLOOR': Math.floor,
  'FLOOR.MATH': Math.floor,
  'ABS': Math.abs,
  'SQRT': Math.sqrt,
  'VLOOKUP': vlookup,
  'MAX': max,
  'SUM': sum,
  'MIN': min,
  'CONCATENATE': concatenate,
  'IF': _if,
  'PMT': pmt,
  'COUNTA': counta,
  'IRR': irr,
  'NORM.INV': normsInv,
  '_xlfn.NORM.INV': normsInv,
  'STDEV': stDeviation,
  'AVERAGE': avg,
  'EXP': EXP,
  'LN': Math.log,
  '_xlfn.VAR.P': var_p,
  'VAR.P': var_p,
  '_xlfn.COVARIANCE.P': covariance_p,
  'COVARIANCE.P': covariance_p,
  'TRIM': trim,
  'LEN': len,
};

// +---------------------+
// | THE IMPLEMENTATIONS |
// +---------------------+

function len(a) {
  return ('' + a).length;
}

function trim(a) {
  return ('' + a).trim();
}

function covariance_p(a, b) {
  const _a = getArrayOfNumbers(a);
  const _b = getArrayOfNumbers(b);
  if (_a.length !== _b.length) {
    return 'N/D';
  }
  const inv_n = 1.0 / _a.length;
  const avg_a = sum(..._a) / _a.length;
  const avg_b = sum(..._b) / _b.length;
  let s = 0.0;
  for (let i = 0; i < _a.length; i ++) {
    s += (_a[i] - avg_a) * (_b[i] - avg_b);
  }
  return s * inv_n;
}

function getArrayOfNumbers(range) {
  const arr = [];
  for (const arg of range) {
    if (Array.isArray(arg)) {
      const matrix = arg;
      for (let j = matrix.length; j--;) {
        if (typeof(matrix[j]) === 'number') {
          arr.push(matrix[j]);
        } else if (Array.isArray(matrix[j])) {
          for (let k = matrix[j].length; k--;) {
            if (typeof(matrix[j][k]) === 'number') {
              arr.push(matrix[j][k]);
            }
          }
        }
        // else {
        //   wtf is that?
        // }
      }
    } else {
      if (typeof(arg) === 'number') {
        arr.push(arg);
      }
    }
  }
  return arr;
}

function var_p(...args) {
  const average = avg(...args);
  let s = 0.0;
  let c = 0;
  for (const arg of args) {
    if (Array.isArray(arg)) {
      const matrix = arg;
      for (let j = matrix.length; j--;) {
        for (let k = matrix[j].length; k--;) {
          if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
            s += Math.pow(matrix[j][k] - average, 2);
            c++;
          }
        }
      }
    } else {
      s += Math.pow(arg - average, 2);
      c++;
    }
  }
  return s / c;
}

function EXP(n) {
  return Math.pow(Math.E, n);
}

function avg(...args) {
  return sum(...args) / counta(...args);
}

function stDeviation(...args) {
  const array = getArrayOfNumbers(args);
  function _mean(_array) {
    return _array.reduce(function(a, b) {
      return a + b;
    }) / _array.length;
  }
  const mean = _mean(array);
  const dev = array.map(function(itm) {
    return (itm - mean) * (itm - mean);
  });
  return Math.sqrt(dev.reduce(function(a, b) {
    return a + b;
  }) / (array.length - 1));
}

/// Original C++ implementation found at http://www.wilmott.com/messageview.cfm?catid=10&threadid=38771
/// C# implementation found at http://weblogs.asp.net/esanchez/archive/2010/07/29/a-quick-and-dirty-implementation-of-excel-norminv-function-in-c.aspx
/*
  *     Compute the quantile function for the normal distribution.
  *
  *     For small to moderate probabilities, algorithm referenced
  *     below is used to obtain an initial approximation which is
  *     polished with a final Newton step.
  *
  *     For very large arguments, an algorithm of Wichura is used.
  *
  *  REFERENCE
  *
  *     Beasley, J. D. and S. G. Springer (1977).
  *     Algorithm AS 111: The percentage points of the normal distribution,
  *     Applied Statistics, 26, 118-121.
  *
  *      Wichura, M.J. (1988).
  *      Algorithm AS 241: The Percentage Points of the Normal Distribution.
  *      Applied Statistics, 37, 477-484.
  */
function normsInv(p, mu, sigma) {
  if (p < 0 || p > 1) {
    throw new Error('The probality p must be bigger than 0 and smaller than 1');
  }
  if (sigma < 0) {
    throw new Error('The standard deviation sigma must be positive');
  }

  if (p === 0) {
    return -Infinity;
  }
  if (p === 1) {
    return Infinity;
  }
  if (sigma === 0) {
    return mu;
  }

  let q;
  let r;
  let val;

  q = p - 0.5;

  /*-- use AS 241 --- */
  /* double ppnd16_(double *p, long *ifault)*/
  /*      ALGORITHM AS241  APPL. STATIST. (1988) VOL. 37, NO. 3
          Produces the normal deviate Z corresponding to a given lower
          tail area of P; Z is accurate to about 1 part in 10**16.
  */
  if (Math.abs(q) <= .425) { /* 0.075 <= p <= 0.925 */
    r = .180625 - q * q;
    val =
      q * (((((((r * 2509.0809287301226727 +
              33430.575583588128105) * r + 67265.770927008700853) * r +
            45921.953931549871457) * r + 13731.693765509461125) * r +
          1971.5909503065514427) * r + 133.14166789178437745) * r +
        3.387132872796366608) / (((((((r * 5226.495278852854561 +
            28729.085735721942674) * r + 39307.89580009271061) * r +
          21213.794301586595867) * r + 5394.1960214247511077) * r +
        687.1870074920579083) * r + 42.313330701600911252) * r + 1);
  } else { /* closer than 0.075 from {0,1} boundary */

    /* r = min(p, 1-p) < 0.075 */
    if (q > 0)
      r = 1 - p;
    else
      r = p;

    r = Math.sqrt(-Math.log(r));
    /* r = sqrt(-log(r))  <==>  min(p, 1-p) = exp( - r^2 ) */

    if (r <= 5) { /* <==> min(p,1-p) >= exp(-25) ~= 1.3888e-11 */
      r += -1.6;
      val = (((((((r * 7.7454501427834140764e-4 +
                .0227238449892691845833) * r + .24178072517745061177) *
              r + 1.27045825245236838258) * r +
            3.64784832476320460504) * r + 5.7694972214606914055) *
          r + 4.6303378461565452959) * r +
        1.42343711074968357734) / (((((((r *
                1.05075007164441684324e-9 + 5.475938084995344946e-4) *
              r + .0151986665636164571966) * r +
            .14810397642748007459) * r + .68976733498510000455) *
          r + 1.6763848301838038494) * r +
        2.05319162663775882187) * r + 1);
    } else { /* very close to  0 or 1 */
      r += -5;
      val = (((((((r * 2.01033439929228813265e-7 +
                2.71155556874348757815e-5) * r +
              .0012426609473880784386) * r + .026532189526576123093) *
            r + .29656057182850489123) * r +
          1.7848265399172913358) * r + 5.4637849111641143699) *
        r + 6.6579046435011037772) / (((((((r *
              2.04426310338993978564e-15 + 1.4215117583164458887e-7) *
            r + 1.8463183175100546818e-5) * r +
          7.868691311456132591e-4) * r + .0148753612908506148525) * r + .13692988092273580531) * r +
        .59983220655588793769) * r + 1);
    }

    if (q < 0.0) {
      val = -val;
    }
  }

  return mu + sigma * val;
}

function irr(range, guess) {
  let _min = -2.0;
  let _max = 1.0;
  let n = 0;
  let NPV;
  let guest;
  do {
    guest = (_min + _max) / 2;
    NPV = 0;
    for (let i = 0; i < range.length; i++) {
      const arg = range[i];
      NPV += arg[0] / Math.pow((1 + guest), i);
    }
    if (NPV > 0) {
      if (_min === _max) {
        _max += Math.abs(guest);
      }
      _min = guest;
    } else {
      _max = guest;
    }
    n++;
  } while (Math.abs(NPV) > 0.000001 && n < 100000);

  return guest;
}

function counta(...args) {
  let r = 0;
  for (const arg of args) {
    if (Array.isArray(arg)) {
      const matrix = arg;
      for (let j = matrix.length; j--;) {
        for (let k = matrix[j].length; k--;) {
          if (matrix[j][k] !== null && matrix[j][k] !== undefined) {
            r++;
          }
        }
      }
    } else {
      if (arg !== null && arg !== undefined) {
        r++;
      }
    }
  }
  return r;
}

function pmt(rate_per_period, number_of_payments, present_value, future_value = 0, type = 0) {
  if (rate_per_period !== 0.0) {
    // Interest rate exists
    const q = Math.pow(1 + rate_per_period, number_of_payments);
    return -(rate_per_period * (future_value + (q * present_value))) / ((-1 + q) * (1 + rate_per_period * (type)));

  } else if (number_of_payments !== 0.0) {
    // No interest rate, but number of payments exists
    return -(future_value + present_value) / number_of_payments;
  }
  return 0;
}

function _if(condition, _then, _else) {
  if (condition) {
    return _then;
  } else {
    return _else;
  }
}

function concatenate(...args) {
  let r = '';
  for (const arg of args) {
    r += arg;
  }
  return r;
}

function sum(...args) {
  let r = 0;
  for (const arg of args) {
    if (Array.isArray(arg)) {
      const matrix = arg;
      for (let j = matrix.length; j--;) {
        for (let k = matrix[j].length; k--;) {
          r += +matrix[j][k];
        }
      }
    } else {
      r += +arg;
    }
  }
  return r;
}

function max(...args) {
  let _max = null;
  for (const arg of args) {
    if (Array.isArray(arg)) {
      const arr = arg;
      for (let j = arr.length; j--;) {
        _max = _max == null || _max < arr[j] ? arr[j] : _max;
      }
    } else if (!isNaN(arg)) {
      _max = _max == null || _max < arg ? arg : _max;
    } else {
      console.log('WTF??', arg);
    }
  }
  return _max;
}

function min(...args) {
  let result = null;
  for (const arg of args) {
    if (Array.isArray(arg)) {
      const arr = arg;
      for (let j = arr.length; j--;) {
        result = result == null || result > arr[j] ? arr[j] : result;
      }
    } else if (!isNaN(arg)) {
      result = result == null || result > arg ? arg : result;
    } else {
      console.log('WTF??', arg);
    }
  }
  return result;
}

function vlookup(key, matrix, return_index) {
  for (const row of matrix) {
    if (row[0] === key) {
      return row[return_index - 1];
    }
  }
  throw Error('#N/A');
}

function my_assign(dest, source) {
  const obj = JSON.parse(JSON.stringify(dest));
  const keys = Object.keys(source);
  for (const k of keys) {
    obj[k] = source[k];
  }
  return obj;
}

class UserFnExecutor {
  private name;
  private args;

  constructor(private user_function) {
    this.name = 'UserFn';
    this.args = [];
  }

  public calc() {
    return this.user_function.apply(this, this.args);
  }

  public push(buffer) {
    this.args.push(buffer.calc());
  }
}

class RawValue {
  constructor(private value) {}

  public calc() {
    return this.value;
  }
}

class RefValue {
  constructor(private str_expression, private formula) {}

  public calc() {
    let cell_name;
    let sheet;
    let sheet_name;
    if (this.str_expression.indexOf('!') !== -1) {
      const aux = this.str_expression.split('!');
      sheet = this.formula.wb.Sheets[aux[0]];
      if (!sheet) {
        const quoted = aux[0].match(/^'(.*)'$/);
        if (quoted) {
          aux[0] = quoted[1];
        }
        sheet = this.formula.wb.Sheets[aux[0]];
      }
      sheet_name = aux[0];
      cell_name = aux[1];
    } else {
      sheet = this.formula.sheet;
      sheet_name = this.formula.sheet_name;
      cell_name = this.str_expression;
    }
    const cell_full_name = sheet_name + '!' + cell_name;
    const ref_cell = sheet[cell_name];
    if (!ref_cell) {
      throw Error('Cell ' + cell_full_name + ' not found.');
    }
    const formula_ref = this.formula.formula_ref[cell_full_name];
    if (formula_ref) {
      if (formula_ref.status === 'new') {
        exec_formula(formula_ref);
        return sheet[cell_name].v;
      } else if (formula_ref.status === 'working') {
        throw new Error('Circular ref');
      } else if (formula_ref.status === 'done') {
        return sheet[cell_name].v;
      }
    } else {
      return sheet[cell_name].v;
    }
  }
}

class Range {
  constructor(private str_expression, private formula) {}

  public calc() {
    let range_expression;
    let sheet_name;
    let sheet;
    if (this.str_expression.indexOf('!') !== -1) {
      const aux = this.str_expression.split('!');
      sheet_name = aux[0];
      range_expression = aux[1];
    } else {
      sheet_name = this.formula.sheet_name;
      range_expression = this.str_expression;
    }
    sheet = this.formula.wb.Sheets[sheet_name];
    const arr = range_expression.split(':');
    const min_row = parseInt(arr[0].replace(/^[A-Z]+/, ''), 10) || 0;
    let str_max_row = arr[1].replace(/^[A-Z]+/, '');
    let max_row;
    if (str_max_row === '' && sheet['!ref']) {
      str_max_row = sheet['!ref'].split(':')[1].replace(/^[A-Z]+/, '');
    }
    // the max is 1048576, but TLE
    max_row = parseInt(str_max_row === '' ? '500000' : str_max_row, 10);
    const min_col = col_str_2_int(arr[0]);
    const max_col = col_str_2_int(arr[1]);
    const matrix = [];
    for (let i = min_row; i <= max_row; i++) {
      const row = [];
      matrix.push(row);
      for (let j = min_col; j <= max_col; j++) {
        const cell_name = int_2_col_str(j) + i;
        const cell_full_name = sheet_name + '!' + cell_name;
        if (this.formula.formula_ref[cell_full_name]) {
          if (this.formula.formula_ref[cell_full_name].status === 'new') {
            exec_formula(this.formula.formula_ref[cell_full_name]);
          } else if (this.formula.formula_ref[cell_full_name].status === 'working') {
            throw new Error('Circular ref');
          }
          row.push(sheet[cell_name].v);
        } else if (sheet[cell_name]) {
          row.push(sheet[cell_name].v);
        } else {
          row.push(null);
        }
      }
    }
    return matrix;
  }
}

let exp_id = 0;

class Exp {
  private id: number;
  private args: any[];
  private name: string;
  private last_arg;

  constructor(private formula) {
    this.id = ++exp_id;
    this.args = [];
    this.name = 'Expression';
  }

  public calc() {
    this.exec_minus();
    this.exec('^', function(a, b) {
      return Math.pow(+a, +b);
    });
    this.exec('*', function(a, b) {
      return (+a) * (+b);
    });
    this.exec('/', function(a, b) {
      return (+a) / (+b);
    });
    this.exec('+', function(a, b) {
      return (+a) + (+b);
    });
    this.exec('&', function(a, b) {
      return '' + a + b;
    });
    this.exec('<', function(a, b) {
      return a < b;
    });
    this.exec('>', function(a, b) {
      return a > b;
    });
    this.exec('>=', function(a, b) {
      return a >= b;
    });
    this.exec('<=', function(a, b) {
      return a <= b;
    });
    this.exec('<>', function(a, b) {
      return a !== b;
    });
    this.exec('=', function(a, b) {
      return a === b;
    });
    if (this.args.length === 1) {
      if (typeof(this.args[0].calc) !== 'function') {
        return this.args[0];
      } else {
        return this.args[0].calc();
      }
    }
  }

  public push(buffer) {
    if (buffer) {
      let v;
      if (!isNaN(buffer)) {
        v = new RawValue(+buffer);
      } else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), this.formula);
      } else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+:[A-Z]+[0-9]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), this.formula);
      } else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+:[A-Z]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), this.formula);
      } else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+:[A-Z]+$/)) {
        v = new Range(buffer.trim().replace(/\$/g, ''), this.formula);
      } else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[A-Z]+[0-9]+$/)) {
        v = new RefValue(buffer.trim().replace(/\$/g, ''), this.formula);
      } else if (typeof buffer === 'string' && buffer.trim().replace(/\$/g, '').match(/^[^!]+![A-Z]+[0-9]+$/)) {
        v = new RefValue(buffer.trim().replace(/\$/g, ''), this.formula);
      } else if (typeof buffer === 'string' && !isNaN(Number(buffer.trim().replace(/%$/, '')))) {
        v = new RawValue(+(buffer.trim().replace(/%$/, '')) / 100.0);
      } else {
        v = buffer;
      }
      if (((v === '=') && (this.last_arg === '>' || this.last_arg === '<')) || (this.last_arg === '<' && v === '>')) {
        this.args[this.args.length - 1] += v;
      } else {
        this.args.push(v);
      }
      this.last_arg = v;
    }
  }

  private exec(op, fn) {
    for (let i = 0; i < this.args.length; i++) {
      if (this.args[i] === op) {
        try {
          const r = fn(this.args[i - 1].calc(), this.args[i + 1].calc());
          this.args.splice(i - 1, 3, new RawValue(r));
          i--;
        } catch (e) {
          throw Error(this.formula.name + ': evaluating ' + this.formula.cell.f + '\n' + e.message);
        }
      }
    }
  }

  private exec_minus() {
    for (let i = this.args.length; i--;) {
      if (this.args[i] === '-') {
        const r = -this.args[i + 1].calc();
        if (typeof this.args[i - 1] !== 'string' && i > 0) {
          this.args.splice(i, 1, '+');
          this.args.splice(i + 1, 1, new RawValue(r));
        } else {
          this.args.splice(i, 2, new RawValue(r));
        }
      }
    }
  }
}

const common_operations = {
  '*': 'multiply',
  '+': 'plus',
  '-': 'minus',
  '/': 'divide',
  '^': 'power',
  '&': 'concat',
  '<': 'lt',
  '>': 'gt',
  '=': 'eq',
};

interface Fn {
  exp: any;
  special?: any;
}

function exec_formula(formula) {
  formula.status = 'working';
  let root_exp;
  let str_formula = formula.cell.f;
  if (str_formula[0] === '=') {
    str_formula = str_formula.substr(1);
  }
  let exp_obj = root_exp = new Exp(formula);
  let buffer = '';
  let is_string = false;
  let was_string = false;
  const fn_stack: Fn[] = [{
    exp: exp_obj,
  }];
  for (const token of str_formula) {
    if (token === '"') {
      if (is_string) {
        exp_obj.push(new RawValue(buffer));
        is_string = false;
        was_string = true;
      } else {
        is_string = true;
      }
      buffer = '';
    } else if (is_string) {
      buffer += token;
    } else if (token === '(') {
      let o;
      const trim_buffer = buffer.trim();
      let special = xlsx_Fx[trim_buffer];
      if (special) {
        // special = new UserFnExecutor(special, formula);
        special = new UserFnExecutor(special /* formula */);
      } else if (trim_buffer) {
        // Error: "Worksheet 1"!D145: Function INDEX not found
        throw new Error('"' + formula.sheet_name + '"!' + formula.name + ': Function ' + buffer + ' not found');
      }
      o = new Exp(formula);
      fn_stack.push({
        exp: o,
        special,
      });
      exp_obj = o;
      buffer = '';
    } else if (common_operations[token]) {
      if (!was_string) {
        exp_obj.push(buffer);
      }
      was_string = false;
      exp_obj.push(token);
      buffer = '';
    } else if (token === ',' && fn_stack[fn_stack.length - 1].special) {
      was_string = false;
      fn_stack[fn_stack.length - 1].exp.push(buffer);
      fn_stack[fn_stack.length - 1].special.push(fn_stack[fn_stack.length - 1].exp);
      fn_stack[fn_stack.length - 1].exp = exp_obj = new Exp(formula);
      buffer = '';
    } else if (token === ')') {
      let v;
      const stack = fn_stack.pop();
      exp_obj = stack.exp;
      exp_obj.push(buffer);
      v = exp_obj;
      buffer = '';
      exp_obj = fn_stack[fn_stack.length - 1].exp;
      if (stack.special) {
        stack.special.push(v);
        exp_obj.push(stack.special);
      } else {
        exp_obj.push(v);
      }
    } else {
      buffer += token;
    }
  }
  root_exp.push(buffer);
  try {
    formula.cell.v = root_exp.calc();
    if (typeof(formula.cell.v) === 'string') {
      formula.cell.t = 's';
    } else if (typeof(formula.cell.v) === 'number') {
      formula.cell.t = 'n';
    }
  } catch (e) {
    if (e.message === '#N/A') {
      formula.cell.v = 42;
      formula.cell.t = 'e';
      formula.cell.w = e.message;
    } else {
      throw e;
    }
  } finally {
    formula.status = 'done';
  }
}

function find_all_cells_with_formulas(wb) {
  const formula_ref = {};
  const cells = [];
  const keys = Object.keys(wb.Sheets);
  for (const sheet_name of keys) {
    const sheet = wb.Sheets[sheet_name];
    for (const cell_name in sheet) {
      if (sheet[cell_name].f) {
        const formula = formula_ref[sheet_name + '!' + cell_name] = {
          formula_ref,
          wb,
          sheet,
          sheet_name,
          cell: sheet[cell_name],
          name: cell_name,
          status: 'new',
        };
        cells.push(formula);
      }
    }
  }
  return cells;
}

export function col_str_2_int(col_str) {
  let r = 0;
  const colstr = col_str.replace(/[0-9]+$/, '');
  for (let i = colstr.length; i--;) {
    r += Math.pow(26, colstr.length - i - 1) * (colstr.charCodeAt(i) - 64);
  }
  return r - 1;
}

export function int_2_col_str(n) {
  let dividend = n + 1;
  let columnName = '';
  let modulo;
  let guard = 10;
  while (dividend > 0 && guard --) {
      modulo = (dividend - 1) % 26;
      columnName = String.fromCharCode(modulo + 65) + columnName;
      dividend = (dividend - modulo - 1) / 26;
  }
  return columnName;
}

export function set_fx(name, fn) {
  xlsx_Fx[name] = fn;
}

export function exec_fx(name, args) {
  return xlsx_Fx[name].apply(this, args);
}

export function import_functions(formulajs, opts: any = {}) {
  const prefix = opts.prefix || '';
  const keys = Object.keys(formulajs);
  for (const key of keys) {
    const obj = formulajs[key];
    if (typeof(obj) === 'function') {
      xlsx_Fx[prefix + key] = obj;
    } else if (typeof(obj) === 'object') {
      import_functions(obj, my_assign(opts, {prefix: key + '.'}));
    }
  }
}

/**
 * Immutable invalidate function
 */
export function invalidate(workbook: WorkBook) {
  const next = cloneDeep(workbook);
  _invalidate(next);
  return next;
}

export default function _invalidate(workbook: WorkBook) {
  const formulas = find_all_cells_with_formulas(workbook);
  for (let i = formulas.length - 1; i >= 0; i--) {
    exec_formula(formulas[i]);
  }
}
