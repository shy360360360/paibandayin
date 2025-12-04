<!--
 * 生成单据打印页面
-->
<script setup>
import { ref, computed, onMounted, watch, nextTick } from 'vue';
import { bitable } from '@lark-base-open/js-sdk';
import { Setting } from '@element-plus/icons-vue';

// --- 工具函数 ---

/**
 * 创建持久化引用
 * @param {string} key - localStorage键名
 * @param {*} defaultValue - 默认值
 * @returns {Object} 响应式引用对象
 */
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

  // 监听数据变化并保存到localStorage
  watch(dataRef, (newValue) => {
    if (Array.isArray(newValue) || (typeof newValue === 'object' && newValue !== null)) {
      localStorage.setItem(key, JSON.stringify(newValue));
    } else {
      localStorage.setItem(key, String(newValue));
    }
  }, { deep: true });

  return dataRef;
}

/**
 * 格式化单元格数据
 * @param {*} cellValue - 单元格原始值
 * @returns {string} 格式化后的字符串
 */
function formatCell(cellValue) {
  if (!cellValue) return '';

  // 处理时间戳格式的日期
  if (typeof cellValue === 'number' && String(cellValue).length === 13) {
    const date = new Date(cellValue);
    if (!isNaN(date.getTime())) {
      const year = date.getFullYear();
      const month = (date.getMonth() + 1).toString().padStart(2, '0');
      const day = date.getDate().toString().padStart(2, '0');
      return `${year}-${month}-${day}`;
    }
  }

  // 处理数组类型
  if (Array.isArray(cellValue)) {
    return cellValue.map(item => {
      if (typeof item === 'object' && item !== null) {
        return item.text || item.name || '';
      }
      return item;
    }).join(', ');
  }

  // 处理对象类型
  if (typeof cellValue === 'object' && cellValue !== null) {
    return cellValue.text || '';
  }

  // 处理基本类型
  return String(cellValue);
}

/**
 * 调度场算法实现表达式计算
 * @param {string} expression - 数学表达式字符串
 * @returns {number} 计算结果
 */
function shuntingYardEvaluate(expression) {
  const outputQueue = [];
  const operatorStack = [];

  // 运算符优先级
  const precedence = {
    '+': 1,
    '-': 1,
    '*': 2,
    '/': 2
  };

  // 运算符结合性
  const associativity = {
    '+': 'L',
    '-': 'L',
    '*': 'L',
    '/': 'L'
  };

  // 解析表达式中的数字和运算符
  const tokens = expression.match(/\d+\.?\d*|[+\-*/()]/g) || [];

  for (let i = 0; i < tokens.length; i++) {
    const token = tokens[i];

    // 数字直接入队
    if (/^\d+\.?\d*$/.test(token)) {
      outputQueue.push(parseFloat(token));
    }
    // 左括号入栈
    else if (token === '(') {
      operatorStack.push(token);
    }
    // 右括号，弹出运算符直到遇到左括号
    else if (token === ')') {
      while (operatorStack.length > 0 && operatorStack[operatorStack.length - 1] !== '(') {
        outputQueue.push(operatorStack.pop());
      }
      if (operatorStack.length === 0) {
        throw new Error('括号不匹配');
      }
      operatorStack.pop(); // 弹出左括号
    }
    // 运算符处理
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

  // 弹出栈中剩余的运算符
  while (operatorStack.length > 0) {
    const op = operatorStack.pop();
    if (op === '(' || op === ')') {
      throw new Error('括号不匹配');
    }
    outputQueue.push(op);
  }

  // 计算后缀表达式
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

/**
 * 根据格式配置格式化数字
 * @param {number} value - 要格式化的数值
 * @param {string} formatConfig - 格式配置（如"2"表示保留2位小数四舍五入，"2入"表示保留2位小数向上取整）
 * @returns {string} 格式化后的数字字符串
 */
function formatNumberByConfig(value, formatConfig) {
  if (!formatConfig) return value;

  let decimalPlaces = 0;
  let roundMode = 'round'; // 默认四舍五入

  // 解析格式配置
  if (formatConfig.match(/^\d+$/)) {
    // 只有数字，表示保留小数位数，使用四舍五入
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
      // 向上取整
      if (value < 0) {
        result = Math.floor(value * multiplier) / multiplier;
      } else {
        result = Math.ceil(value * multiplier) / multiplier;
      }
      break;
    case '舍':
      // 向下取整
      if (value < 0) {
        result = Math.ceil(value * multiplier) / multiplier;
      } else {
        result = Math.floor(value * multiplier) / multiplier;
      }
      break;
    default:
      // 四舍五入
      result = Math.round(value * multiplier) / multiplier;
  }

  // 确保显示指定的小数位数
  return result.toFixed(decimalPlaces);
}

/**
 * 将数字转换为中文大写金额格式
 * @param {number} number - 要转换的数字
 * @returns {string} 中文大写金额字符串
 */
function convertToChineseAmount(number) {
  // 中文数字
  const chineseNumbers = ['零', '壹', '贰', '叁', '肆', '伍', '陆', '柒', '捌', '玖'];

  // 中文单位
  const chineseUnits = ['', '拾', '佰', '仟'];
  const chineseBigUnits = ['', '万', '亿'];

  // 处理负数
  const isNegative = number < 0;
  number = Math.abs(number);

  // 保留两位小数
  const roundedNumber = Math.round(number * 100) / 100;
  const [integerPart, decimalPart = ''] = roundedNumber.toString().split('.');

  // 处理整数部分
  let integerStr = integerPart;
  let result = '';

  // 补齐到能被4整除的位数
  while (integerStr.length % 4 !== 0) {
    integerStr = '0' + integerStr;
  }

  // 按4位一组处理
  const groups = [];
  for (let i = 0; i < integerStr.length; i += 4) {
    groups.push(integerStr.substring(i, i + 4));
  }

  // 处理每一组
  for (let i = 0; i < groups.length; i++) {
    const group = groups[i];
    let groupStr = '';
    let zeroFlag = false;

    for (let j = 0; j < group.length; j++) {
      const digit = parseInt(group[j]);

      if (digit === 0) {
        // 处理零的情况
        if (groupStr !== '' && !zeroFlag) {
          groupStr += '零';
          zeroFlag = true;
        }
      } else {
        // 处理非零数字
        groupStr += chineseNumbers[digit] + chineseUnits[3 - j];
        zeroFlag = false;
      }
    }

    // 去掉末尾的零
    groupStr = groupStr.replace(/零+$/, '');

    // 添加大单位
    if (groupStr !== '') {
      groupStr += chineseBigUnits[groups.length - 1 - i];
    }

    result += groupStr;
  }

  // 处理连续的零
  result = result.replace(/零+/g, '零');

  // 特殊情况：只有零
  if (result === '') {
    result = '零';
  }

  // 添加元
  result += '元';

  // 处理小数部分
  if (decimalPart) {
    const decimalDigits = decimalPart.padEnd(2, '0').substring(0, 2);
    const jiao = parseInt(decimalDigits[0]);
    const fen = parseInt(decimalDigits[1]);

    if (jiao > 0) {
      result += chineseNumbers[jiao] + '角';
    } else if (fen > 0) {
      // 如果有分但没有角，需要加零
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

  // 处理负数
  if (isNegative) {
    result = '负' + result;
  }

  return result;
}

/**
 * 变量替换函数
 * @param {string} text - 包含变量的文本
 * @param {Object} customerData - 客户数据
 * @param {Object} calculationResults - 计算结果
 * @param {Array} allFields - 所有字段信息
 * @param {Array} groupingFieldIds - 分组字段ID数组
 * @returns {string} 替换变量后的文本
 */
function replaceVariables(text, customerData, calculationResults, allFields, groupingFieldIds) {
  try {
    if (!text || !customerData) return text;

    // 处理\n和\t的重复语法
    let processedText = text
      .replace(/\\n\*(\d+)/g, (match, count) => '<br>'.repeat(parseInt(count)))
      .replace(/\\t\*(\d+)/g, (match, count) => '&nbsp;'.repeat(parseInt(count)))
      .replace(/\\n/g, '<br>')
      .replace(/\\t/g, '&nbsp;');

    // 处理被@符号包裹的{{变量名}}引用，转换为中文大写金额格式
    processedText = processedText.replace(/@(\{\{[^}]+\}\})@/g, (match, variable) => {
      // 先处理变量引用，获取实际值
      let value = variable;

      // 处理带索引的计算结果引用
      value = value.replace(/\{\{\s*([^}\s\[\]]+)\s*\[\s*(\d+)\s*\]\s*\}\}/g, (match, varName, indexStr) => {
        const trimmedVarName = varName.trim();
        const index = parseInt(indexStr, 10);

        // 检查计算结果是否存在且索引有效
        if (calculationResults &&
          calculationResults[trimmedVarName] &&
          calculationResults[trimmedVarName].values &&
          Array.isArray(calculationResults[trimmedVarName].values) &&
          index >= 0 &&
          index < calculationResults[trimmedVarName].values.length &&
          calculationResults[trimmedVarName].values[index] !== undefined) {

          const value = calculationResults[trimmedVarName].values[index];
          const formatConfig = calculationResults[trimmedVarName].formatConfig;
          // 根据格式配置格式化结果
          return formatNumberByConfig(value, formatConfig);
        }

        return match; // 未找到对应计算结果，保留原始文本
      });

      // 处理简化形式的计算结果引用
      value = value.replace(/\{\{\s*([^}\s\[\]]+)\s*\}\}/g, (match, varName) => {
        const trimmedVarName = varName.trim();

        // 检查是否存在计算结果且不为字段名
        if (calculationResults &&
          calculationResults[trimmedVarName] &&
          calculationResults[trimmedVarName].values &&
          Array.isArray(calculationResults[trimmedVarName].values) &&
          calculationResults[trimmedVarName].values.length > 0 &&
          calculationResults[trimmedVarName].values[0] !== undefined &&
          !allFields.find(f => f.name === trimmedVarName)) {

          const value = calculationResults[trimmedVarName].values[0];
          const formatConfig = calculationResults[trimmedVarName].formatConfig;
          // 根据格式配置格式化结果
          return formatNumberByConfig(value, formatConfig);
        }

        return match; // 保留原始匹配，让下一个正则处理
      });

      // 处理普通字段引用
      value = value.replace(/\{\{([^}]*)\}\}/g, (match, fieldName) => {
        const trimmedFieldName = fieldName.trim();
        const field = allFields.find(f => f.name === trimmedFieldName);
        if (!field) return match;

        try {
          if (groupingFieldIds.length === 0) {
            // 非分组模式，从第一条记录获取
            const record = customerData.records[0];
            if (record && typeof record[field.id] !== 'undefined') {
              return formatCell(record[field.id]);
            }
          } else {
            // 分组模式，从分组信息获取
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

      // 如果最终结果是数字，则转换为中文大写金额
      const numericResult = parseFloat(value);
      if (!isNaN(numericResult)) {
        return convertToChineseAmount(numericResult);
      }

      return value;
    });

    // 处理被@符号包裹的数字，转换为中文大写金额格式
    processedText = processedText.replace(/@(\d+\.?\d*)@/g, (match, number) => {
      const num = parseFloat(number);
      if (!isNaN(num)) {
        return convertToChineseAmount(num);
      }
      return match;
    });

    // 处理带索引的计算结果引用
    processedText = processedText.replace(/\{\{\s*([^}\s\[\]]+)\s*\[\s*(\d+)\s*\]\s*\}\}/g, (match, varName, indexStr) => {
      const trimmedVarName = varName.trim();
      const index = parseInt(indexStr, 10);

      // 检查计算结果是否存在且索引有效
      if (calculationResults &&
        calculationResults[trimmedVarName] &&
        calculationResults[trimmedVarName].values &&
        Array.isArray(calculationResults[trimmedVarName].values) &&
        index >= 0 &&
        index < calculationResults[trimmedVarName].values.length &&
        calculationResults[trimmedVarName].values[index] !== undefined) {

        const value = calculationResults[trimmedVarName].values[index];
        const formatConfig = calculationResults[trimmedVarName].formatConfig;
        // 根据格式配置格式化结果
        const result = formatNumberByConfig(value, formatConfig);
        return result;
      }

      return match; // 未找到对应计算结果，保留原始文本
    });

    // 处理简化形式的计算结果引用
    processedText = processedText.replace(/\{\{\s*([^}\s\[\]]+)\s*\}\}/g, (match, varName) => {
      const trimmedVarName = varName.trim();

      // 检查是否存在计算结果且不为字段名
      if (calculationResults &&
        calculationResults[trimmedVarName] &&
        calculationResults[trimmedVarName].values &&
        Array.isArray(calculationResults[trimmedVarName].values) &&
        calculationResults[trimmedVarName].values.length > 0 &&
        calculationResults[trimmedVarName].values[0] !== undefined &&
        !allFields.find(f => f.name === trimmedVarName)) {

        const value = calculationResults[trimmedVarName].values[0];
        const formatConfig = calculationResults[trimmedVarName].formatConfig;
        // 根据格式配置格式化结果
        const result = formatNumberByConfig(value, formatConfig);
        return result;
      }

      return match; // 保留原始匹配，让下一个正则处理
    });

    // 处理普通字段引用
    processedText = processedText.replace(/\{\{([^}]*)\}\}/g, (match, fieldName) => {
      const trimmedFieldName = fieldName.trim();
      const field = allFields.find(f => f.name === trimmedFieldName);
      if (!field) return match;

      try {
        if (groupingFieldIds.length === 0) {
          // 非分组模式，从第一条记录获取
          const record = customerData.records[0];
          if (record && typeof record[field.id] !== 'undefined') {
            const result = formatCell(record[field.id]);
            return result;
          }
        } else {
          // 分组模式，从分组信息获取
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

    const lines = processedText.split('<br>');
    const newLines = lines.map(line => {
      const rightParts = [];
      const leftLine = line.replace(/>([^>]+?)>/g, (match, content) => {
        rightParts.push(content);
        return '';
      });

      if (rightParts.length > 0) {
        return `<div style="display: flex; justify-content: space-between; align-items: center; width: 100%;"><span>${leftLine}</span><span>${rightParts.join('')}</span></div>`;
      }
      return line;
    });
    processedText = newLines.join('<br>');

    return processedText;
  } catch (error) {
    console.error('replaceVariables函数出错:', error);
    return text;
  }
}

// --- 状态管理 ---

// 控制预览对话框显示状态
const showInvoice = ref(false);
// 控制打印设置对话框显示状态
const showSettingsDialog = ref(false);
// 控制全部打印对话框显示状态
const showAllInvoiceDialog = ref(false);
// 存储所有字段信息
const allFields = ref([]);
const invoiceDataByCustomer = ref([]);
const currentCustomerIndex = ref(0);
let lastRecordIdList = [];

// --- 持久化状态 ---
// 选择显示的字段ID数组
const selectedFieldIds = createPersistentRef('selectedFieldIds', []);
// 分组合并依据字段ID数组
const groupingFieldIds = createPersistentRef('groupingFieldIds', []);
// 存储列宽配置
const columnWidths = createPersistentRef('columnWidths', '');
// 左上角文本内容
const topLeftText = createPersistentRef('topLeftText', '');
// 右上角文本内容
const topRightText = createPersistentRef('topRightText', '');
// 左下角文本内容
const bottomLeftText = createPersistentRef('bottomLeftText', '');
// 右下角文本内容
const bottomRightText = createPersistentRef('bottomRightText', '');
// 发票标题
const invoiceTitle = createPersistentRef('invoiceTitle', '');
// 标题字体大小
const titleFontSize = createPersistentRef('titleFontSize', 28);
// 标题是否加粗
const isTitleBold = createPersistentRef('isTitleBold', true);
// 打印宽度缩放百分比
const printScaleWidth = createPersistentRef('printScaleWidth', 100);
// 打印高度缩放百分比
const printScaleHeight = createPersistentRef('printScaleHeight', 100);
// 是否保持宽高比
const keepAspectRatio = createPersistentRef('keepAspectRatio', true);
// 打印上边距
const printMarginTop = createPersistentRef('printMarginTop', 0);
// 打印左边距
const printMarginLeft = createPersistentRef('printMarginLeft', 0);
// 表格后空白行数
const emptyRowsAfterTable = createPersistentRef('emptyRowsAfterTable', 0);
// 是否合并最后一行
const isLastRowMerged = createPersistentRef('isLastRowMerged', false);
// 合并行文本内容
const mergedRowText = createPersistentRef('mergedRowText', '');
// 计算公式配置
const calculationFormulas = createPersistentRef('calculationFormulas', '');
// 表格内容居中
const isTableDataCentered = createPersistentRef('isTableDataCentered', false);
// 合并格内容居中
const isMergedCellCentered = createPersistentRef('isMergedCellCentered', false);
// 计算结果存储
const calculationResults = ref({});

// --- 计算属性 ---

// 根据 selectedFieldIds 计算出实际要显示的字段对象，并保持选择顺序
const displayedFields = computed(() => {
  if (!allFields.value.length) return [];
  return selectedFieldIds.value.map(id => allFields.value.find(f => f.id === id)).filter(Boolean);
});

const columnWidthStyles = computed(() => {
  if (!columnWidths.value || displayedFields.value.length === 0) {
    return displayedFields.value.map(() => ({}));
  }

  const widths = columnWidths.value.split(',').map(w => parseFloat(w.trim()) || 0);
  const totalProportion = widths.reduce((sum, w) => sum + w, 0);

  if (totalProportion === 0) {
    return displayedFields.value.map(() => ({}));
  }

  return displayedFields.value.map((field, index) => {
    const width = widths[index];
    if (width > 0) {
      const percentage = (width / totalProportion) * 100;
      return { width: `${percentage}%` };
    } else {
      return {}; // '0' or invalid numbers result in dynamic width
    }
  });
});

const titleStyle = computed(() => ({
  fontSize: `${titleFontSize.value}px`,
  fontWeight: isTitleBold.value ? 'bold' : 'normal',
}));

// --- 计算文本内容 ---

/**
 * 获取客户文本内容
 * @param {Object} customerData - 客户数据
 * @returns {Object} 包含四个角落文本的对象
 */
function getCustomerText(customerData) {
  const result = {
    topLeft: replaceVariables(topLeftText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value),
    topRight: replaceVariables(topRightText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value),
    bottomLeft: replaceVariables(bottomLeftText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value),
    bottomRight: replaceVariables(bottomRightText.value, customerData, calculationResults.value, allFields.value, groupingFieldIds.value)
  };
  return result;
}

// 左上角文本计算属性
const computedTopLeftText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return topLeftText.value;
  }
  return replaceVariables(topLeftText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});

// 右上角文本计算属性
const computedTopRightText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return topRightText.value;
  }
  return replaceVariables(topRightText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});

// 左下角文本计算属性
const computedBottomLeftText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return bottomLeftText.value;
  }
  return replaceVariables(bottomLeftText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});

// 右下角文本计算属性
const computedBottomRightText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return bottomRightText.value;
  }
  return replaceVariables(bottomRightText.value, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});

// 合并行文本计算属性
const computedMergedRowText = computed(() => {
  if (invoiceDataByCustomer.value.length === 0 || currentCustomerIndex.value >= invoiceDataByCustomer.value.length) {
    return mergedRowText.value || '&nbsp;';
  }
  // 如果配置了合并行文本，则使用配置的文本，否则使用默认的空白字符
  const textToProcess = mergedRowText.value || '&nbsp;';
  return replaceVariables(textToProcess, invoiceDataByCustomer.value[currentCustomerIndex.value], calculationResults.value, allFields.value, groupingFieldIds.value);
});

// --- 计算逻辑处理 ---

/**
 * 处理计算逻辑
 * @param {Object} customerData - 客户数据
 */
function processCalculations(customerData) {
  if (!calculationFormulas.value || !customerData || !customerData.records.length) {
    calculationResults.value = {};
    return;
  }

  const results = {};
  const formulas = calculationFormulas.value.split(';').map(f => f.trim()).filter(f => f);

  formulas.forEach(formula => {
    // 解析公式定义
    const match = formula.match(/^([^=()]+(?:\([^)]+\))?)\s*=\s*(.+)$/);
    if (!match) {
      console.error('公式格式错误:', formula);
      return;
    }

    const varDef = match[1].trim();
    const expression = match[2].trim();

    // 解析变量名和格式设置
    let varName = varDef;
    let formatConfig = null;

    // 检查是否有格式设置
    const formatMatch = varDef.match(/^([^(]+)\(([^)]+)\)$/);
    if (formatMatch) {
      varName = formatMatch[1].trim();
      formatConfig = formatMatch[2].trim();
    }

    // 处理特殊聚合操作：字段名+
    if (expression.endsWith('+') && !expression.includes('+ ')) {
      const fieldName = expression.slice(0, -1).trim();
      const field = allFields.value.find(f => f.name === fieldName);
      if (field) {
        let sum = 0;
        customerData.records.forEach(record => {
          const value = parseFloat(formatCell(record[field.id])) || 0;
          sum += value;
        });
        // 存储格式配置和结果
        results[varName] = {
          values: [sum],
          formatConfig: formatConfig
        };
      }
      return;
    }

    // 处理特殊聚合操作：字段名*
    if (expression.endsWith('*') && !expression.includes('* ')) {
      const fieldName = expression.slice(0, -1).trim();
      const field = allFields.value.find(f => f.name === fieldName);
      if (field) {
        let product = 1;
        customerData.records.forEach(record => {
          const value = parseFloat(formatCell(record[field.id])) || 1;
          product *= value;
        });
        // 存储格式配置和结果
        results[varName] = {
          values: [product],
          formatConfig: formatConfig
        };
      }
      return;
    }

    // 处理普通表达式：对每一行数据进行计算
    const rowResults = [];

    // 遍历每条记录，为每条记录计算一次表达式
    customerData.records.forEach((record, index) => {
      try {
        // 创建字段值映射对象
        const fieldValues = {};

        // 先检查表达式中包含哪些字段名
        const expressionFields = [];
        allFields.value.forEach(field => {
          // 使用更宽松的匹配方式，支持中文字段名
          const escapedFieldName = field.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          const fieldRegex = new RegExp(`(?:^|\\s|[\\(\\)+\\-*/])${escapedFieldName}(?:$|\\s|[\\(\\)+\\-*/])`, 'g');
          if (fieldRegex.test(' ' + expression + ' ')) {
            expressionFields.push(field);
          }
        });

        // 遍历表达式中包含的字段，获取当前记录的字段值
        expressionFields.forEach(field => {
          try {
            // 确保字段ID在记录中存在
            if (record && record[field.id] !== undefined) {
              const rawValue = formatCell(record[field.id]);
              // 直接尝试将原始值转换为数字
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

        // 深拷贝表达式，避免修改原始字符串
        let evaluatedExpression = expression;

        // 按字段名长度降序排序，确保较长的字段名先被替换
        const sortedFieldNames = Object.keys(fieldValues).sort((a, b) => b.length - a.length);

        // 替换表达式中的字段名为实际值，使用更精确的替换方式
        sortedFieldNames.forEach(fieldName => {
          if (fieldValues[fieldName] !== undefined) {
            // 使用单词边界匹配的正则表达式，确保只替换完整的字段名
            const escapedFieldName = fieldName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
            const regex = new RegExp(`(^|\\s|[\\(\\)+\\-*/])${escapedFieldName}($|\\s|[\\(\\)+\\-*/])`, 'g');
            evaluatedExpression = evaluatedExpression.replace(regex, `$1${fieldValues[fieldName]}$2`);
          }
        });

        // 使用支持运算符优先级的表达式计算器
        const result = evaluateExpression(evaluatedExpression);

        // 确保结果是数字类型
        const finalResult = Number(result);
        const cleanResult = isNaN(finalResult) ? 0 : finalResult;

        rowResults.push(cleanResult);
      } catch (error) {
        console.error(`记录${index}计算${varName}出错:`, error);
        // 出错时添加0
        rowResults.push(0);
      }
    });

    // 存储格式配置和结果
    results[varName] = {
      values: rowResults,
      formatConfig: formatConfig
    };
  });

  calculationResults.value = results;
}

/**
 * 支持运算符优先级的表达式计算器
 * @param {string} expression - 数学表达式字符串
 * @returns {number} 计算结果
 */
function evaluateExpression(expression) {
  // 移除所有空格
  expression = expression.replace(/\s+/g, '');

  // 验证表达式只包含数字、运算符、小数点和括号
  if (!/^[\d+\-*/().]+$/.test(expression)) {
    throw new Error('表达式包含非法字符');
  }

  try {
    // 使用调度场算法处理运算符优先级
    return shuntingYardEvaluate(expression);
  } catch (error) {
    console.error('表达式计算错误:', error);
    return 0;
  }
}

/**
 * 获取统一的打印样式
 * @returns {string} CSS样式字符串
 */
const getPrintStyles = () => {
  const scaleWidth = printScaleWidth.value / 100;
  const scaleHeight = printScaleHeight.value / 100;

  const transforms = [];

  // 仅在边距不为0时应用位移
  if (printMarginLeft.value !== 0 || printMarginTop.value !== 0) {
    transforms.push(`translate(${printMarginLeft.value}px, ${printMarginTop.value}px)`);
  }

  // 仅在缩放不为100%时应用缩放
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

// --- 生命周期钩子 ---

/**
 * 组件加载时获取字段列表，用于填充多选框
 */
onMounted(async () => {
  try {
    const table = await bitable.base.getActiveTable();
    const fieldMetaList = await table.getFieldMetaList();

    // 加载所有字段，用于选择显示列和合并依据
    allFields.value = fieldMetaList;

    // 设置定时器，每100毫秒检查一次选择变化
    setInterval(loadData, 100);

    // 加载初始数据
    await loadData();

  } catch (error) {
    console.error('Failed to initialize plugin:', error);
  }
});

// --- 监听器 ---

// 监听输入变化并保存到 localStorage
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

// 缩放比例同步更新标志
let isUpdatingScale = false;

// 监听等比缩放选项变化
watch(keepAspectRatio, (newValue) => {
  if (newValue) {
    printScaleHeight.value = printScaleWidth.value;
  }
});

// 监听宽度缩放变化
watch(printScaleWidth, (newValue) => {
  if (keepAspectRatio.value && !isUpdatingScale) {
    isUpdatingScale = true;
    printScaleHeight.value = newValue;
    nextTick(() => { isUpdatingScale = false; });
  }
});

// 监听高度缩放变化
watch(printScaleHeight, (newValue) => {
  if (keepAspectRatio.value && !isUpdatingScale) {
    isUpdatingScale = true;
    printScaleWidth.value = newValue;
    nextTick(() => { isUpdatingScale = false; });
  }
});

// --- 主要功能函数 ---

/**
 * 加载并处理选中数据
 */
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

  // 检查选中记录是否有变化
  if (JSON.stringify(recordIdList) === JSON.stringify(lastRecordIdList)) {
    return; // 没有变化则不重新生成
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

/**
 * 生成发票预览
 */
async function generateInvoice(recordIdList, table) {
  // 移除了"请至少选择一个要显示的列！"的提示，允许用户不选择任何列
  // if (selectedFieldIds.value.length === 0) {
  //   alert('请至少选择一个要显示的列！');
  //   return;
  // }

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
  // 生成发票后处理计算逻辑
  if (invoiceDataByCustomer.value.length > 0) {
    processCalculations(invoiceDataByCustomer.value[currentCustomerIndex.value]);
  }
  showInvoice.value = true;
}

/**
 * 打印当前单据
 */
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

  // 移除分页符，因为只打印一页
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

/**
 * 打印全部单据 (iframe版本)
 */
function printAllInvoices() {
  // 1. 创建一个隐藏的iframe
  const iframe = document.createElement('iframe');
  iframe.style.position = 'absolute';
  iframe.style.width = '0';
  iframe.style.height = '0';
  iframe.style.border = 'none';
  iframe.style.left = '-9999px';
  iframe.style.top = '-9999px';
  document.body.appendChild(iframe);

  const doc = iframe.contentWindow.document;

  // 2. 构建所有发票页面的HTML内容
  let allInvoicesHtml = '';
  invoiceDataByCustomer.value.forEach(customerData => {
    // 为每个客户生成计算结果
    processCalculations(customerData);

    const topLeftHtml = getCustomerText(customerData).topLeft || '&nbsp;';
    const topRightHtml = getCustomerText(customerData).topRight || '&nbsp;';
    const bottomLeftHtml = getCustomerText(customerData).bottomLeft || '&nbsp;';
    const bottomRightHtml = getCustomerText(customerData).bottomRight || '&nbsp;';
    const mergedRowHtml = computedMergedRowText.value || '&nbsp;';

    let recordsHtml = '';
    customerData.records.forEach(record => {
      recordsHtml += '<tr>';
      displayedFields.value.forEach((field, index) => {
        const style = columnWidthStyles.value[index] && columnWidthStyles.value[index].width ? `style="width: ${columnWidthStyles.value[index].width}"` : '';
        recordsHtml += `<td ${style}>${formatCell(record[field.id])}</td>`;
      });
      recordsHtml += '</tr>';
    });

    // 添加空白行
    for (let i = 0; i < emptyRowsAfterTable.value; i++) {
      recordsHtml += '<tr>';
      displayedFields.value.forEach((field, index) => {
        const style = columnWidthStyles.value[index] && columnWidthStyles.value[index].width ? `style="width: ${columnWidthStyles.value[index].width}"` : '';
        recordsHtml += `<td ${style}>&nbsp;</td>`;
      });
      recordsHtml += '</tr>';
    }

    // 添加合并行
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
                ${displayedFields.value.map((field, index) => {
      const style = columnWidthStyles.value[index] && columnWidthStyles.value[index].width ? `style="width: ${columnWidthStyles.value[index].width}"` : '';
      return `<th ${style}>${field.name}</th>`;
    }).join('')}
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

  // 3. 注入HTML和样式到iframe
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

  // 4. 执行打印并清理
  iframe.contentWindow.focus();
  setTimeout(() => {
    iframe.contentWindow.print();
    document.body.removeChild(iframe);
  }, 500); // 等待iframe内容渲染
}

/**
 * 切换到上一个客户
 */
function prevCustomer() {
  if (currentCustomerIndex.value > 0) {
    currentCustomerIndex.value--;
    // 切换客户后重新处理计算逻辑
    processCalculations(invoiceDataByCustomer.value[currentCustomerIndex.value]);
  }
}

/**
 * 切换到下一个客户
 */
function nextCustomer() {
  if (currentCustomerIndex.value < invoiceDataByCustomer.value.length - 1) {
    currentCustomerIndex.value++;
    // 切换客户后重新处理计算逻辑
    processCalculations(invoiceDataByCustomer.value[currentCustomerIndex.value]);
  }
}
</script>

<template>
  <div class="container">
    <!-- 设置按钮 -->
    <div class="settings-bar">
      <el-button :icon="Setting" circle @click="showSettingsDialog = true"></el-button>
    </div>

    <!-- 嵌入式打印预览 -->
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
                <th v-for="(field, index) in displayedFields" :key="field.id" :style="columnWidthStyles[index]">{{
                  field.name }}</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(record, idx) in invoiceDataByCustomer[currentCustomerIndex].records" :key="idx">
                <td v-for="(field, index) in displayedFields" :key="field.id"
                  :style="[{ textAlign: isTableDataCentered ? 'center' : 'left' }, columnWidthStyles[index]]">
                  {{ formatCell(record[field.id]) }}
                </td>
              </tr>
              <!-- 表格下方空白行 -->
              <tr v-for="rowIndex in emptyRowsAfterTable" :key="'empty-' + rowIndex">
                <td v-for="(field, index) in displayedFields" :key="field.id" :style="columnWidthStyles[index]">&nbsp;
                </td>
              </tr>
              <!-- 根据设置决定是否添加合并的最后一行 -->
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

    <!-- 设置对话框 -->
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
          <el-form-item label="显示列列宽">
            <el-input v-model="columnWidths" placeholder="用逗号分隔,如 7,6,5。0表示动态宽度"></el-input>
          </el-form-item>
          <el-form-item label="合并依据">
            <el-select v-model="groupingFieldIds" multiple placeholder="请选择用于合并单据的列" style="width: 100%;">
              <el-option v-for="field in allFields" :key="field.id" :label="field.name" :value="field.id">
              </el-option>
            </el-select>
          </el-form-item>

          <!-- 左上角文本 -->
          <el-form-item label="左上角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="topLeftText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额，>内容> 表示内容靠右展示">
                </el-input>
              </div>
            </div>
          </el-form-item>

          <!-- 右上角文本 -->
          <el-form-item label="右上角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="topRightText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额，>内容> 表示内容靠右展示">
                </el-input>
              </div>
            </div>
          </el-form-item>

          <!-- 左下角文本 -->
          <el-form-item label="左下角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="bottomLeftText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额，>内容> 表示内容靠右展示">
                </el-input>
              </div>
            </div>
          </el-form-item>

          <!-- 右下角文本 -->
          <el-form-item label="右下角文本">
            <div style="display: flex; flex-direction: column; width: 100%;">
              <div style="flex-grow: 1;">
                <el-input v-model="bottomRightText" type="textarea" :rows="2"
                  placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额，>内容> 表示内容靠右展示">
                </el-input>
              </div>
            </div>
          </el-form-item>

          <!-- 计算逻辑配置 -->
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

            <!-- 合并行文本配置项 -->
            <el-form-item label="合并行文本" v-show="isLastRowMerged">
              <div style="display: flex; flex-direction: column; width: 100%;">
                <div style="flex-grow: 1;">
                  <el-input v-model="mergedRowText" type="textarea" :rows="2"
                    placeholder="支持使用{{字段名}}引用表格数据，\n表示换行（可*n叠加），\t表示空格（可*n叠加），@数值@ 表示金额转换为中文大写金额，>内容> 表示内容靠右展示"></el-input>
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
  table-layout: fixed;
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