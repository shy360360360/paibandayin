<!--
 * 生成单据打印页面
-->
<script setup>
import { ref, computed, onMounted, watch, nextTick } from 'vue';
import { bitable } from '@lark-base-open/js-sdk';
import { Setting } from '@element-plus/icons-vue';



function createPersistentRef(key, defaultValue) {
  const storedValue = localStorage.getItem(key);
  let initialValue;

  if (storedValue === null) {
    initialValue = defaultValue;
  } else {
    if (typeof defaultValue === 'boolean') {
      initialValue = storedValue === 'true';
    } else if (typeof defaultValue === 'number') {
      initialValue = Number(storedValue);
    } else if (Array.isArray(defaultValue) || (typeof defaultValue === 'object' && defaultValue !== null)) {
      try {
        initialValue = JSON.parse(storedValue);
      } catch {
        initialValue = defaultValue;
      }
    } else {
      initialValue = storedValue;
    }
  }

  const dataRef = ref(initialValue);


  watch(dataRef, (newValue) => {
    if (Array.isArray(newValue) || (typeof newValue === 'object' && newValue !== null)) {
      localStorage.setItem(key, JSON.stringify(newValue));
    } else {
      localStorage.setItem(key, String(newValue));
    }
  }, { deep: true });

  return dataRef;
}


function formatCell(cellValue) {
  if (!cellValue) return '';


  if (typeof cellValue === 'number' && String(cellValue).length === 13) {
    const date = new Date(cellValue);
    if (!isNaN(date.getTime())) {
      const year = date.getFullYear();
      const month = (date.getMonth() + 1).toString().padStart(2, '0');
      const day = date.getDate().toString().padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  }


  if (Array.isArray(cellValue)) {
    return cellValue.map(item => {
      if (typeof item === 'object' && item !== null) {
        return item.text || item.name || '';
      }
      return item;
    }).join(', ');
  }


  if (typeof cellValue === 'object' && cellValue !== null) {
    return cellValue.text || '';
  }


  return String(cellValue);
}


function shuntingYardEvaluate(expression) {
  const outputQueue = [];
  const operatorStack = [];

  const precedence = {
    '+': 1,
    '-': 1,
    '*': 2,
    '/': 2
  };

  const associativity = {
    '+': 'L',
    '-': 'L',
    '*': 'L',
    '/': 'L'
  };

  const tokens = expression.match(/\d+\.?\d*|[+\-*/()]/g) || [];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    if (/^\d+\.?\d*$/.test(token)) {
      outputQueue.push(parseFloat(token));
    }
    else if (token === '(') {
      operatorStack.push(token);
    }
    else if (token === ')') {
      while (operatorStack.length > 0 && operatorStack[operatorStack.length - 1] !== '(') {
        outputQueue.push(operatorStack.pop());
      }
      if (operatorStack.length === 0) {
        throw new Error('括号不匹配');
      }
      operatorStack.pop();
    }
    else if (token in precedence) {
      while (
        operatorStack.length > 0 &&
        operatorStack[operatorStack.length - 1] !== '(' &&
        precedence[operatorStack[operatorStack.length - 1]] !== undefined &&
        (
          precedence[operatorStack[operatorStack.length - 1]] > precedence[token] ||
          (precedence[operatorStack[operatorStack.length - 1]] === precedence[token] && associativity[token] === 'L')
        )
      ) {
        outputQueue.push(operatorStack.pop());
      }
      operatorStack.push(token);
    }
  }

  while (operatorStack.length > 0) {
    const op = operatorStack.pop();
    if (op === '(' || op === ')') {
      throw new Error('括号不匹配');
    }
    outputQueue.push(op);
  }

  const stack = [];
  for (let i = 0; i < outputQueue.length; i++) {
    const token = outputQueue[i];

    if (typeof token === 'number') {
      stack.push(token);
    } else {
      if (stack.length < 2) {
        throw new Error('表达式格式错误');
      }

      const b = stack.pop();
      const a = stack.pop();

      switch (token) {
        case '+':
          stack.push(a + b);
          break;
        case '-':
          stack.push(a - b);
          break;
        case '*':
          stack.push(a * b);
          break;
        case '/':
          if (b === 0) {
            throw new Error('除零错误');
          }
          stack.push(a / b);
          break;
        default:
          throw new Error(`未知运算符: ${token}`);
      }
    }
  }

  if (stack.length !== 1) {
    throw new Error('表达式格式错误');
  }

  return stack[0];
}


function formatNumberByConfig(value, formatConfig) {
  if (!formatConfig) return value;

  let decimalPlaces = 0;
  let roundMode = 'round';

  if (formatConfig.match(/^\d+$/)) {
    decimalPlaces = parseInt(formatConfig, 10);
  } else if (formatConfig.match(/^(\d+)(入|舍)$/)) {
    const match = formatConfig.match(/^(\d+)(入|舍)$/);
    decimalPlaces = parseInt(match[1], 10);
    roundMode = match[2];
  }

  const multiplier = Math.pow(10, decimalPlaces);
  let result;

  switch (roundMode) {
    case '入':
      if (value < 0) {
        result = Math.floor(value * multiplier) / multiplier;
      } else {
        result = Math.ceil(value * multiplier) / multiplier;
      }
      break;
    case '舍':
      if (value < 0) {
        result = Math.ceil(value * multiplier) / multiplier;
      } else {
        result = Math.floor(value * multiplier) / multiplier;
      }
      break;
    default:
      result = Math.round(value * multiplier) / multiplier;
  }

  return result.toFixed(decimalPlaces);
}


function convertToChineseAmount(number) {
  const chineseNumbers = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];

  const chineseUnits = ['', '拾', '佰', '仟'];
  const chineseBigUnits = ['', '万', '亿'];

  const isNegative = number < 0;
  number = Math.abs(number);

  const roundedNumber = Math.round(number * 100) / 100;
  const [integerPart, decimalPart = ''] = roundedNumber.toString().split('.');

  let integerStr = integerPart;
  let result = '';

  while (integerStr.length % 4 !== 0) {
    integerStr = '0' + integerStr;
  }

  const groups = [];
  for (let i = 0; i < integerStr.length; i += 4) {
    groups.push(integerStr.substring(i, i + 4));
  }

  for (let i = 0; i < groups.length; i++) {
    const group = groups[i];
    let groupStr = '';
    let zeroFlag = false;

    for (let j = 0; j < group.length; j++) {
      const digit = parseInt(group[j]);

      if (digit === 0) {
        if (groupStr !== '' && !zeroFlag) {
          groupStr += '零';
          zeroFlag = true;
        }
      } else {
        groupStr += chineseNumbers[digit] + chineseUnits[3 - j];
        zeroFlag = false;
      }
    }

    groupStr = groupStr.replace(/零+$/, '');

    if (groupStr !== '') {
      groupStr += chineseBigUnits[groups.length - 1 - i];
    }

    result += groupStr;
  }

  result = result.replace(/零+/g, '零');

  if (result === '') {
    result = '零';
  }

  result += '元';

  if (decimalPart) {
    const decimalDigits = decimalPart.padEnd(2, '0').substring(0, 2);
    const jiao = parseInt(decimalDigits[0]);
    const fen = parseInt(decimalDigits[1]);

    if (jiao > 0) {
      result += chineseNumbers[jiao] + '角';
    } else if (fen > 0) {
      if (jiao === 0) {
        result += '零';
      }
      result += chineseNumbers[fen] + '分';
    } else if (jiao === 0 && fen === 0) {
      result += '整';
    }

    if (jiao > 0 && fen > 0) {
      result += chineseNumbers[fen] + '分';
    } else if (jiao > 0 && fen === 0) {
      result += '整';
    }
  } else {
    result += '整';
  }

  if (isNegative) {
    result = '负' + result;
  }

  return result;
}


function replaceVariables(text, customerData, calculationResults, allFields, groupingFieldIds) {
  try {
    if (!text || !customerData) return text;

    let processedText = text
      .replace(/\\n\*(\d+)/g, (match, count) => '<br>'.repeat(parseInt(count)))
      .replace(/\\t\*(\d+)/g, (match, count) => '&nbsp;'.repeat(parseInt(count)))
      .replace(/\\n/g, '<br>')
      .replace(/\\t/g, '&nbsp;');


    processedText = processedText.replace(/@(\{\{[^}]+\}\})@/g, (match, variable) => {
      let value = variable;


      value = value.replace(/\{\{\s*([^}\s\[\]]+)\s*\[\s*(\d+)\s*\]\s*\}\}/g, (match, varName, indexStr) => {
        const trimmedVarName = varName.trim();
        const index = parseInt(indexStr, 10);


        if (calculationResults &&
          calculationResults[trimmedVarName] &&
          calculationResults[trimmedVarName].values &&
          Array.isArray(calculationResults[trimmedVarName].values) &&
          index >= 0 &&
          index < calculationResults[trimmedVarName].values.length &&
          calculationResults[trimmedVarName].values[index] !== undefined) {

          const value = calculationResults[trimmedVarName].values[index];
          const formatConfig = calculationResults[trimmedVarName].formatConfig;
          return formatNumberByConfig(value, formatConfig);
        }

        return match;
      });


      value = value.replace(/\{\{\s*([^}\s\[\]]+)\s*\}\}/g, (match, varName) => {
        const trimmedVarName = varName.trim();


        if (calculationResults &&
          calculationResults[trimmedVarName] &&
          calculationResults[trimmedVarName].values &&
          Array.isArray(calculationResults[trimmedVarName].values) &&
          calculationResults[trimmedVarName].values.length > 0 &&
          calculationResults[trimmedVarName].values[0] !== undefined &&
          !allFields.find(f => f.name === trimmedVarName)) {

          const value = calculationResults[trimmedVarName].values[0];
          const formatConfig = calculationResults[trimmedVarName].formatConfig;
          return formatNumberByConfig(value, formatConfig);
        }

        return match;
      });


      value = value.replace(/\{\{([^}]*)\}\}/g, (match, fieldName) => {
        const trimmedFieldName = fieldName.trim();
        const field = allFields.find(f => f.name === trimmedFieldName);
        if (!field) return match;

        try {
          if (groupingFieldIds.length === 0) {
            const record = customerData.records[0];
            if (record && typeof record[field.id] !== 'undefined') {
              return formatCell(record[field.id]);
            }
          } else {
            const groupValue = customerData.groupingInfo.find(g => g.name === field.name);
            if (groupValue) {
              return groupValue.value;
            }
          }
        } catch (error) {
          console.error('替换普通字段出错:', error);
        }
        return '';
      });


      const numericResult = parseFloat(value);
      if (!isNaN(numericResult)) {
        return convertToChineseAmount(numericResult);
      }

      return value;
    });


    processedText = processedText.replace(/@(\d+\.?\d*)@/g, (match, number) => {
      const num = parseFloat(number);
      if (!isNaN(num)) {
        return convertToChineseAmount(num);
      }
      return match;
    });


    processedText = processedText.replace(/\{\{\s*([^}\s\[\]]+)\s*\[\s*(\d+)\s*\]\s*\}\}/g, (match, varName, indexStr) => {
      const trimmedVarName = varName.trim();
      const index = parseInt(indexStr, 10);


      if (calculationResults &&
        calculationResults[trimmedVarName] &&
        calculationResults[trimmedVarName].values &&
        Array.isArray(calculationResults[trimmedVarName].values) &&
        index >= 0 &&
        index < calculationResults[trimmedVarName].values.length &&
        calculationResults[trimmedVarName].values[index] !== undefined) {

        const value = calculationResults[trimmedVarName].values[index];
        const formatConfig = calculationResults[trimmedVarName].formatConfig;
        const result = formatNumberByConfig(value, formatConfig);
        return result;
      }

      return match;
    });


    processedText = processedText.replace(/\{\{\s*([^}\s\[\]]+)\s*\}\}/g, (match, varName) => {
      const trimmedVarName = varName.trim();


      if (calculationResults &&
        calculationResults[trimmedVarName] &&
        calculationResults[trimmedVarName].values &&
        Array.isArray(calculationResults[trimmedVarName].values) &&
        calculationResults[trimmedVarName].values.length > 0 &&
        calculationResults[trimmedVarName].values[0] !== undefined &&
        !allFields.find(f => f.name === trimmedVarName)) {

        const value = calculationResults[trimmedVarName].values[0];
        const formatConfig = calculationResults[trimmedVarName].formatConfig;
        const result = formatNumberByConfig(value, formatConfig);
        return result;
      }

      return match;
    });


    processedText = processedText.replace(/\{\{([^}]*)\}\}/g, (match, fieldName) => {
      const trimmedFieldName = fieldName.trim();
      const field = allFields.find(f => f.name === trimmedFieldName);
      if (!field) return match;

      try {
        if (groupingFieldIds.length === 0) {
          const record = customerData.records[0];
          if (record && typeof record[field.id] !== 'undefined') {
            const result = formatCell(record[field.id]);
            return result;
          }
        } else {
          const groupValue = customerData.groupingInfo.find(g => g.name === field.name);
          if (groupValue) {
            const result = groupValue.value;
            return result;
          }
        }
      } catch (error) {
        console.error('替换普通字段出错:', error);
      }
      return '';
    });

    return processedText;
  } catch (error) {
    console.error('replaceVariables函数出错:', error);
    return text;
  }
}



const showInvoice = ref(false);
const showSettingsDialog = ref(false);
const showAllInvoiceDialog = ref(false);
const allFields = ref([]);
const invoiceDataByCustomer = ref([]);
const currentCustomerIndex = ref(0);
let lastRecordIdList = [];


const selectedFieldIds = createPersistentRef('selectedFieldIds', []);
const groupingFieldIds = createPersistentRef('groupingFieldIds', []);
const topLeftText = createPersistentRef('topLeftText', '');
const topRightText = createPersistentRef('topRightText', '');
const bottomLeftText = createPersistentRef('bottomLeftText', '');
const bottomRightText = createPersistentRef('bottomRightText', '');
const invoiceTitle = createPersistentRef('invoiceTitle', '');
const titleFontSize = createPersistentRef('titleFontSize', 28);
const isTitleBold = createPersistentRef('isTitleBold', true);
const printScaleWidth = createPersistentRef('printScaleWidth', 100);
const printScaleHeight = createPersistentRef('printScaleHeight', 100);
const keepAspectRatio = createPersistentRef('keepAspectRatio', true);
const printMarginTop = createPersistentRef('printMarginTop', 0);
const printMarginLeft = createPersistentRef('printMarginLeft', 0);
const emptyRowsAfterTable = createPersistentRef('emptyRowsAfterTable', 0);
const isLastRowMerged = createPersistentRef('isLastRowMerged', false);
const mergedRowText = createPersistentRef('mergedRowText', '');
const calculationFormulas = createPersistentRef('calculationFormulas', '');
const isTableDataCentered = createPersistentRef('isTableDataCentered', false);
const isMergedCellCentered = createPersistentRef('isMergedCellCentered', false);
const calculationResults = ref({});



const displayedFields = computed(() => {
  if (!selectedFieldIds.value.length || !allFields.value.length) {
    return [];
  }
  return selectedFieldIds.value.map(id => {
    return allFields.value.find(field => field.id === id);
  }).filter(Boolean);
});


const titleStyle = computed(() => ({
  fontSize: `${titleFontSize.value}px`,
  fontWeight: isTitleBold.value ? 'bold' : 'normal',
}));



function getCustomerText(customerData) {
  const result = {
    topLeft: replaceVariables(topLeftText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value),
    topRight: replaceVariables(topRightText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value),
    bottomLeft: replaceVariables(bottomLeftText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value),
    bottomRight: replaceVariables(bottomRightText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value)
  };
  return result;
}


const computedTopLeftText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return topLeftText.value;
  }
  return replaceVariables(topLeftText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});


const computedTopRightText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return topRightText.value;
  }
  return replaceVariables(topRightText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});


const computedBottomLeftText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return bottomLeftText.value;
  }
  return replaceVariables(bottomLeftText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});


const computedBottomRightText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return bottomRightText.value;
  }
  return replaceVariables(bottomRightText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});


const computedMergedRowText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return mergedRowText.value || '&nbsp;';
  }
  const textToProcess = mergedRowText.value || '&nbsp;';
  return replaceVariables(textToProcess, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});



function processCalculations(customerData) {
  if (!calculationFormulas.value || !customerData || !customerData.records.length) {
    calculationResults.value = {};
    return;
  }

  const results = {};
  const formulas = calculationFormulas.value.split(';').map(f => f.trim()).filter(f => f);

  formulas.forEach(formula => {
    const match = formula.match(/^([^=()]+(?:\([^)]+\))?)\s*=\s*(.+)$/);
    if (!match) {
      console.error('公式格式错误:', formula);
      return;
    }

    const varDef = match[1].trim();
    const expression = match[2].trim();

    let varName = varDef;
    let formatConfig = null;

    const formatMatch = varDef.match(/^([^(]+)\(([^)]+)\)$/);
    if (formatMatch) {
      varName = formatMatch[1].trim();
      formatConfig = formatMatch[2].trim();
    }

    if (expression.endsWith('+') && !expression.includes('+ ')) {
      const fieldName = expression.slice(0, -1).trim();
      const field = allFields.value.find(f => f.name === fieldName);
      if (field) {
        let sum = 0;
        customerData.records.forEach(record => {
          const value = parseFloat(formatCell(record[field.id])) || 0;
          sum += value;
        });
        results[varName] = {
          values: [sum],
          formatConfig: formatConfig
        };
      }
      return;
    }

    if (expression.endsWith('*') && !expression.includes('* ')) {
      const fieldName = expression.slice(0, -1).trim();
      const field = allFields.value.find(f => f.name === fieldName);
      if (field) {
        let product = 1;
        customerData.records.forEach(record => {
          const value = parseFloat(formatCell(record[field.id])) || 1;
          product *= value;
        });
        results[varName] = {
          values: [product],
          formatConfig: formatConfig
        };
      }
      return;
    }

    const rowResults = [];

    customerData.records.forEach((record, index) => {
      try {
        const fieldValues = {};

        const expressionFields = [];
        allFields.value.forEach(field => {
          const escapedFieldName = field.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const fieldRegex = new RegExp(`(?:^|\\s|[\\(\\)+\\-*/])${escapedFieldName}(?:$|\\s|[\\(\\)+\\-*/])`, 'g');
          if (fieldRegex.test(' ' + expression + ' ')) {
            expressionFields.push(field);
          }
        });

        expressionFields.forEach(field => {
          try {
            if (record && record[field.id] !== undefined) {
              const rawValue = formatCell(record[field.id]);
              const numValue = parseFloat(rawValue);
              fieldValues[field.name] = isNaN(numValue) ? 0 : numValue;
            } else {
              fieldValues[field.name] = 0;
            }
          } catch (e) {
            console.error(`获取记录${index}的${field.name}值出错:`, e);
            fieldValues[field.name] = 0;
          }
        });

        let evaluatedExpression = expression;

        const sortedFieldNames = Object.keys(fieldValues).sort((a, b) => b.length - a.length);

        sortedFieldNames.forEach(fieldName => {
          if (fieldValues[fieldName] !== undefined) {
            const escapedFieldName = fieldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const regex = new RegExp(`(^|\\s|[\\(\\)+\\-*/])${escapedFieldName}($|\\s|[\\(\\)+\\-*/])`, 'g');
            evaluatedExpression = evaluatedExpression.replace(regex, `$1${fieldValues[fieldName]}$2`);
          }
        });

        const result = evaluateExpression(evaluatedExpression);

        const finalResult = Number(result);
        const cleanResult = isNaN(finalResult) ? 0 : finalResult;

        rowResults.push(cleanResult);
      } catch (error) {
        console.error(`记录${index}计算${varName}出错:`, error);
        rowResults.push(0);
      }
    });

    results[varName] = {
      values: rowResults,
      formatConfig: formatConfig
    };
  });

  calculationResults.value = results;
}


function evaluateExpression(expression) {
  expression = expression.replace(/\s+/g, '');

  if (!/^[\d+\-*/().]+$/.test(expression)) {
    throw new Error('表达式包含非法字符');
  }

  try {
    return shuntingYardEvaluate(expression);
  } catch (error) {
    console.error('表达式计算错误:', error);
    return 0;
  }
}


const getPrintStyles = () => {
  const scaleWidth = printScaleWidth.value / 100;
  const scaleHeight = printScaleHeight.value / 100;

  const transforms = [];

  if (printMarginLeft.value !== 0 || printMarginTop.value !== 0) {
    transforms.push(`translate(${printMarginLeft.value}px, ${printMarginTop.value}px)`);
  }

  if (printScaleWidth.value !== 100 || printScaleHeight.value !== 100) {
    transforms.push(`scale(${scaleWidth}, ${scaleHeight})`);
  }

  const transformStyle = transforms.join(' ');

  let styles = `
  body { 
    margin: 0; 
    -webkit-print-color-adjust: exact; 
    print-color-adjust: exact; 
    transform-origin: top;
    transform: ${transformStyle};
  }
  @page { 
    size: A4 portrait; 
    margin: 20mm;
  }
  .invoice-page { 
    page-break-after: always;
    break-after: page;
  }
  .invoice-page:last-child { 
    page-break-after: auto; 
    break-after: auto;
  }
  .invoice-wrapper {
    padding: 24px;
    background-color: #fff;
  }
  .invoice-title {
    text-align: center;
    margin-bottom: 24px;
  }
  .meta-info {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 8px;
    font-size: 16px;
    padding: 0 8px;
    flex-wrap: wrap;
  }
  .invoice-footer {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-top: 10px;
    font-size: 16px;
    padding: 0 8px;
  }
   .invoice-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 16px;
  }
  .invoice-table th,
  .invoice-table td {
    border: 1px solid #000;
    padding: 8px;
  }
  .invoice-table td {
    text-align: ${isTableDataCentered.value ? 'center' : 'left'};
  }
  .invoice-table th {
    font-weight: bold;
    background-color: #f8f9fa;
    text-align: center;
  }
  .invoice-table .merged-row td {
    text-align: ${isMergedCellCentered.value ? 'center' : 'left'};
  }
`;
  return styles;
};



onMounted(async () => {
  try {
    const table = await bitable.base.getActiveTable();
    const fieldMetaList = await table.getFieldMetaList();

    allFields.value = fieldMetaList;

    setInterval(loadData, 100);

    await loadData();

  } catch (error) {
    console.error('Failed to initialize plugin:', error);
  }
});




watch([topLeftText, topRightText, bottomLeftText, bottomRightText, invoiceTitle, titleFontSize, isTitleBold],
  ([newTopLeft, newTopRight, newBottomLeft, newBottomRight, newTitle, newFontSize, newIsBold]) => {
    localStorage.setItem('topLeftText', newTopLeft);
    localStorage.setItem('topRightText', newTopRight);
    localStorage.setItem('bottomLeftText', newBottomLeft);
    localStorage.setItem('bottomRightText', newBottomRight);
    localStorage.setItem('invoiceTitle', newTitle);
    localStorage.setItem('titleFontSize', String(newFontSize));
    localStorage.setItem('isTitleBold', String(newIsBold));
  }
);


let isUpdatingScale = false;


watch(keepAspectRatio, (newValue) => {
  if (newValue) {
    printScaleHeight.value = printScaleWidth.value;
  }
});


watch(printScaleWidth, (newValue) => {
  if (keepAspectRatio.value && !isUpdatingScale) {
    isUpdatingScale = true;
    printScaleHeight.value = newValue;
    nextTick(() => { isUpdatingScale = false; });
  }
});


watch(printScaleHeight, (newValue) => {
  if (keepAspectRatio.value && !isUpdatingScale) {
    isUpdatingScale = true;
    printScaleWidth.value = newValue;
    nextTick(() => { isUpdatingScale = false; });
  }
});




async function loadData() {
  const selection = await bitable.base.getSelection();
  if (!selection || !selection.tableId || !selection.viewId) {
    if (lastRecordIdList.length > 0) {
      invoiceDataByCustomer.value = [];
      showInvoice.value = false;
      lastRecordIdList = [];
    }
    return;
  }

  const table = await bitable.base.getTable(selection.tableId);
  const view = await table.getViewById(selection.viewId);
  const recordIdList = await view.getSelectedRecordIdList();

  if (JSON.stringify(recordIdList) === JSON.stringify(lastRecordIdList)) {
    return;
  }

  lastRecordIdList = recordIdList;

  if (!recordIdList || recordIdList.length === 0) {
    invoiceDataByCustomer.value = [];
    showInvoice.value = false;
    currentCustomerIndex.value = 0;
    return;
  }

  await generateInvoice(recordIdList, table);
}


async function generateInvoice(recordIdList, table) {
  const fieldMetaList = await table.getFieldMetaList();

  let processedData = [];

  if (groupingFieldIds.value.length > 0) {
    const groupingFields = groupingFieldIds.value.map(id =>
      fieldMetaList.find(f => f.id === id)
    ).filter(Boolean);

    if (groupingFields.length !== groupingFieldIds.value.length) {
      alert('部分合并依据字段已不存在，请重新选择。');
      return;
    }

    const groupMap = {};
    for (const recordId of recordIdList) {
      const record = await table.getRecordById(recordId);

      const groupValues = groupingFields.map(field => ({
        name: field.name,
        value: formatCell(record.fields[field.id]) || '无'
      }));

      const groupKey = groupValues.map(g => g.value).join('#');

      if (!groupMap[groupKey]) {
        groupMap[groupKey] = {
          groupingInfo: groupValues,
          records: []
        };
      }
      groupMap[groupKey].records.push(record.fields);
    }
    processedData = Object.values(groupMap);
  } else {
    for (const recordId of recordIdList) {
      const record = await table.getRecordById(recordId);

      processedData.push({
        groupingInfo: [],
        records: [record.fields]
      });
    }
  }

  invoiceDataByCustomer.value = processedData;

  if (invoiceDataByCustomer.value.length === 0) {
    showInvoice.value = false;
    return;
  }

  currentCustomerIndex.value = 0;
  if (invoiceDataByCustomer.value.length > 0) {
    processCalculations(invoiceDataByCustomer.value[currentCustomerIndex.value]);
  }
  showInvoice.value = true;
}


function printInvoice() {
  const iframe = document.createElement('iframe');
  iframe.style.position = 'absolute';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = 'none';
  iframe.style.left = '-9999px';
  iframe.style.top = '-9999px';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;

  const printArea = document.getElementById('invoice-print-area');
  if (!printArea) {
    console.error('找不到打印区域');
    document.body.removeChild(iframe);
    return;
  }
  const currentInvoiceHtml = printArea.innerHTML;

  const styles = getPrintStyles()
    .replace(/page-break-after: always;/g, '')
    .replace(/break-after: page;/g, '');

  const html = `
    <html>
      <head>
        <title>打印预览</title>
        <style>${styles}</style>
      </head>
      <body>
        <div class="invoice-wrapper">
          ${currentInvoiceHtml}
        </div>
      </body>
    </html>
  `;

  doc.open();
  doc.write(html);
  doc.close();

  iframe.contentWindow.focus();
  setTimeout(() => {
    iframe.contentWindow.print();
    document.body.removeChild(iframe);
  }, 500);
}


function printAllInvoices() {
  const iframe = document.createElement('iframe');
  iframe.style.position = 'absolute';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = 'none';
  iframe.style.left = '-9999px';
  iframe.style.top = '-9999px';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;

  let allInvoicesHtml = '';
  invoiceDataByCustomer.value.forEach(customerData => {
    processCalculations(customerData);

    const topLeftHtml = getCustomerText(customerData).topLeft || '&nbsp;';
    const topRightHtml = getCustomerText(customerData).topRight || '&nbsp;';
    const bottomLeftHtml = getCustomerText(customerData).bottomLeft || '&nbsp;';
    const bottomRightHtml = getCustomerText(customerData).bottomRight || '&nbsp;';
    const mergedRowHtml = computedMergedRowText.value || '&nbsp;';

    let recordsHtml = '';
    customerData.records.forEach(record => {
      recordsHtml += '<tr>';
      displayedFields.value.forEach(field => {
        recordsHtml += `<td>${formatCell(record[field.id])}</td>`;
      });
      recordsHtml += '</tr>';
    });

    for (let i = 0; i < emptyRowsAfterTable.value; i++) {
      recordsHtml += '<tr>';
      displayedFields.value.forEach(() => {
        recordsHtml += '<td>&nbsp;</td>';
      });
      recordsHtml += '</tr>';
    }

    if (isLastRowMerged.value) {
      recordsHtml += `
        <tr class="merged-row">
            <td colspan="${displayedFields.value.length}">${mergedRowHtml}</td>
        </tr>
        `;
    }

    allInvoicesHtml += `
      <div class="invoice-page">
        <div class="invoice-wrapper">
          <h1 style="font-size: ${titleFontSize.value}px; font-weight: ${isTitleBold.value ? 'bold' : 'normal'};" class="invoice-title">${invoiceTitle.value}</h1>
          <div class="meta-info">
            <div class="meta-info-top-left">${topLeftHtml}</div>
            <div class="meta-info-top-right">${topRightHtml}</div>
          </div>
          <table class="invoice-table">
            <thead>
              <tr>
                ${displayedFields.value.map(field => `<th>${field.name}</th>`).join('')}
              </tr>
            </thead>
            <tbody>
              ${recordsHtml}
            </tbody>
          </table>
          <div class="invoice-footer">
            <div class="left-footer">${bottomLeftHtml}</div>
            <div class="right-footer">${bottomRightHtml}</div>
          </div>
        </div>
      </div>
    `;
  });

  doc.open();
  doc.write(`
    <html>
      <head>
        <title>打印预览</title>
        <style>
          ${getPrintStyles()}
        </style>
      </head>
      <body>
        ${allInvoicesHtml}
      </body>
    </html>
  `);
  doc.close();

  iframe.contentWindow.focus();
  setTimeout(() => {
    iframe.contentWindow.print();
    document.body.removeChild(iframe);
  }, 500);
}


function prevCustomer() {
  if (currentCustomerIndex.value > 0) {
    currentCustomerIndex.value--;
    processCalculations(invoiceDataByCustomer.value[currentCustomerIndex.value]);
  }
}


function nextCustomer() {
  if (currentCustomerIndex.value < invoiceDataByCustomer.value.length - 1) {
    currentCustomerIndex.value++;
    processCalculations(invoiceDataByCustomer.value[currentCustomerIndex.value]);
  }
}
</script>

<template>
  <div class="container">
    
    <div class="settings-bar">
      <el-button :icon="Setting" circle @click="showSettingsDialog = true"></el-button>
    </div>

    
    <div v-if="showInvoice" class="embedded-preview">
      <div class="invoice-wrapper">
        <div id="invoice-print-area">
          <h1 :style="titleStyle" class="invoice-title">{{ invoiceTitle }}</h1>
          <div class="meta-info">
            <div class="meta-info-top-left" v-html="computedTopLeftText"></div>
            <div class="meta-info-top-right" v-html="computedTopRightText"></div>
          </div>
          <table class="invoice-table">
            <thead>
              <tr>
                <th v-for="field in displayedFields" :key="field.id">{{ field.name }}</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(record, idx) in invoiceDataByCustomer[currentCustomerIndex].records" :key="idx">
                <td v-for="field in displayedFields" :key="field.id"
                  :style="{ textAlign: isTableDataCentered ? 'center' : 'left' }">
                  {{ formatCell(record[field.id]) }}
                </td>
              </tr>
              
              <tr v-for="rowIndex in emptyRowsAfterTable" :key="'empty-' + rowIndex">
                <td v-for="field in displayedFields" :key="field.id">&nbsp;</td>
              </tr>
              
              <tr v-if="isLastRowMerged" :key="'merged-last'" class="merged-row">
                <td :colspan="displayedFields.length" v-html="computedMergedRowText"
                  :style="{ textAlign: isMergedCellCentered ? 'center' : 'left' }"></td>
              </tr>
            </tbody>
          </table>
          <div class="invoice-footer">
            <div class="left-footer" v-html="computedBottomLeftText"></div>
            <div class="right-footer" v-html="computedBottomRightText"></div>
          </div>
        </div>
      </div>
      <div class="preview-actions">
        <el-button @click="prevCustomer" :disabled="currentCustomerIndex === 0">上一张</el-button>
        <span>第 {{ currentCustomerIndex + 1 }} / {{ invoiceDataByCustomer.length }} 张</span>
        <el-button @click="nextCustomer"
          :disabled="currentCustomerIndex === invoiceDataByCustomer.length - 1">下一张</el-button>
        <el-button type="primary" @click="printInvoice">打印当前页</el-button>
        <el-button type="success" @click="printAllInvoices">全部打印</el-button>
      </div>
    </div>
    <div v-else class="empty-state">
      <p>请在表格中选择需要打印的数据行</p>
    </div>

    
    <el-dialog v-model="showSettingsDialog" title="打印设置" width="80%">
      <div style="width: 80%;">
        <el-form label-width="120px">
          <el-form-item label="顶部标题">
            <el-input v-model="invoiceTitle" placeholder="请输入标题"></el-input>
          </el-form-item>
          <el-form-item label="标题大小">
            <div style="display: flex; align-items: center; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input-number v-model="titleFontSize" :min="12" :max="72" controls-position="right"
                  style="width: 100%;">
                </el-input-number>
              </div>
              <el-checkbox v-model="isTitleBold" style="margin-left: 10px; white-space: nowrap;">加粗</el-checkbox>
            </div>
          </el-form-item>
          <el-form-item label="选择显示列">
            <el-select v-model="selectedFieldIds" multiple placeholder="请选择要显示的列" style="width: 100%;">
              <el-option v-for="field in allFields" :key="field.id" :label="field.name" :value="field.id">
              </el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="合并依据">
            <el-select v-model="groupingFieldIds" multiple placeholder="请选择用于合并单据的列" style="width: 100%;">
              <el-option v-for="field in allFields" :key="field.id" :label="field.name" :value="field.id">
              </el-option>
            </el-select>
          </el-form-item>

          
          <el-form-item label="左上角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="topLeftText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额">
                </el-input>
              </div>
            </div>
          </el-form-item>

          
          <el-form-item label="右上角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="topRightText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额">
                </el-input>
              </div>
            </div>
          </el-form-item>

          
          <el-form-item label="左下角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="bottomLeftText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额">
                </el-input>
              </div>
            </div>
          </el-form-item>

          
          <el-form-item label="右下角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="bottomRightText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额">
                </el-input>
              </div>
            </div>
          </el-form-item>

          
          <el-form-item label="计算逻辑配置">
            <el-input v-model="calculationFormulas" type="textarea" :rows="3"
              placeholder="格式：变量名(保留位数保留规则) = 表达式 //保留规则为不填/舍/入，不填为四舍五入;  示例：AAA = 斤数*单价 - 成本；  BBB(2舍) = 金额+ // 金额列求；  CCC = 金额* // 金额列求积；  使用方式：{{AAA[1]}}、{{BBB}}"></el-input>
          </el-form-item>

          <el-form label-width="120px">
            <el-form-item label="宽度缩放">
              <div style="display: flex; align-items: center; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model.number="printScaleWidth" type="number" :min="10" :max="200" :step="10">
                    <template #append>%</template>
                  </el-input>
                </div>
                <el-checkbox v-model="keepAspectRatio" style="margin-left: 10px; white-space: nowrap;">等比</el-checkbox>
              </div>
            </el-form-item>
            <el-form-item label="高度缩放">
              <div style="display: flex; align-items: center; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model.number="printScaleHeight" type="number" :min="10" :max="200" :step="10">
                    <template #append>%</template>
                  </el-input>
                </div>
                <el-checkbox style="margin-left: 10px; visibility: hidden;">占位</el-checkbox>
              </div>
            </el-form-item>
            <el-form-item label="距页面上距离">
              <div style="display: flex; align-items: center; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model.number="printMarginTop" type="number" :max="1000" :step="10">
                    <template #append>px</template>
                  </el-input>
                </div>
                <el-checkbox style="margin-left: 10px; visibility: hidden;">占位</el-checkbox>
              </div>
            </el-form-item>
            <el-form-item label="距页面左距离">
              <div style="display: flex; align-items: center; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model.number="printMarginLeft" type="number" :max="1000" :step="10">
                    <template #append>px</template>
                  </el-input>
                </div>
                <el-checkbox style="margin-left: 10px; visibility: hidden;">占位</el-checkbox>
              </div>
            </el-form-item>
            <el-form-item label="空白行数">
              <div style="display: flex; align-items: center; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model.number="emptyRowsAfterTable" type="number" :max="1000" :step="1">
                    <template #append>行</template>
                  </el-input>
                </div>
                <el-checkbox v-model="isLastRowMerged" style="margin-left: 10px; white-space: nowrap;">合并</el-checkbox>
              </div>
            </el-form-item>

            
            <el-form-item label="合并行文本" v-show="isLastRowMerged">
              <div style="display: flex; flex-direction: column; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model="mergedRowText" type="textarea" :rows="2"
                    placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额"></el-input>
                </div>
              </div>
            </el-form-item>
            <el-form-item label="内容居中">
              <el-checkbox v-model="isTableDataCentered">非合并格</el-checkbox>
              <el-checkbox v-model="isMergedCellCentered">合并格</el-checkbox>
            </el-form-item>
          </el-form>
        </el-form>
      </div>
      <template #footer>
        <span class="dialog-footer">
          <el-button @click="showSettingsDialog = false">关闭</el-button>
        </span>
      </template>
    </el-dialog>
  </div>
</template>

<style>
.container {
  padding: 24px;
}

.settings-bar {
  text-align: center;
  margin-bottom: 20px;
}

.invoice-title {
  text-align: center;
  margin-bottom: 24px;
}

.meta-info {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 8px;
  font-size: 16px;
  padding: 0 8px;
  flex-wrap: wrap;
}

.group-info {
  margin-right: 20px;
  margin-bottom: 8px;
}

.invoice-footer {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-top: 10px;
  font-size: 16px;
  padding: 0 8px;
}

.center-button-container .el-form-item__content {
  justify-content: center;
}

.print-area-wrapper {
  display: flex;
  justify-content: center;
}

.invoice-table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 10px;
}

.invoice-table th,
.invoice-table td {
  border: 1px solid #ddd;
  padding: 8px;
  text-align: left;
}

.invoice-table th {
  background-color: #f2f2f2;
  font-weight: bold;
  text-align: center;
}

.embedded-preview {
  margin-top: 20px;
  border: 1px solid #ebeef5;
  border-radius: 4px;
  overflow: hidden;
}

.invoice-wrapper {
  padding: 24px;
  background-color: #fff;
}

.preview-actions {
  padding: 10px 24px;
  text-align: right;
  border-top: 1px solid #ebeef5;
  background-color: #f5f7fa;
}

.preview-actions span {
  margin: 0 10px;
  color: #606266;
  font-size: 14px;
}

.empty-state {
  margin-top: 20px;
  text-align: center;
  color: #909399;
  padding: 40px;
  border: 1px dashed #dcdfe6;
  border-radius: 4px;
}

@media print {

  .settings-bar,
  .preview-actions,
  .el-dialog,
  .el-overlay {
    display: none !important;
  }
}
</style>