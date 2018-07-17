#!/usr/bin/env node

const path = require('path');
const { writeFileSync, readFileSync, existsSync } = require('fs');
const xlsx = require('node-xlsx');
const Confirm = require('prompt-confirm');
const commandLineArgs = require('command-line-args')
const commandLineUsage = require('command-line-usage')

/**
 * CLI OPTIONS AND USAGE
 */
let options;
const usage = commandLineUsage([
  {
    header: 'EXCEL TO JSON APP',
    content: 'Generate json file from excel file'
  },
  {
    header: 'Options',
    optionList: [
      {
        name: 'createMap',
        typeLabel: '{underline -r}',
        description: 'To create a json file to map the keys'
      },
      {
        name: 'createDict',
        typeLabel: '{underline -k}',
        description: 'To create a json file to translate the keys'
      },
      {
        name: 'file',
        typeLabel: '{underline -f}',
        description: 'Excel file'
      },
      {
        name: 'dictFile',
        typeLabel: '{underline -d}',
        description: 'Json file with key dictionary.'
      },
      {
        name: 'mapFile',
        typeLabel: '{underline -m}',
        description: 'Json file with key map.'
      },
      {
        name: 'keysLine',
        typeLabel: '{underline -l}',
        description: 'Line number (excel) of keys.'
      },
      {
        name: 'sheetIndex',
        typeLabel: '{underline -i}',
        description: 'Index of sheet. (Default: 0)'
      },
      {
        name: 'help',
        typeLabel: '{underline -h}',
        description: 'Show this usage guide.'
      }
    ]
  },
  {
    content: '--file (-f) is default option'
  }
])
try {
  options = commandLineArgs([
    { name: 'createMap', alias: 'r', type: Boolean },
    { name: 'createDict', alias: 'k', type: Boolean },
    { name: 'help', alias: 'h', type: Boolean },
    { name: 'file', alias: 'f', type: String, defaultOption: true },
    { name: 'dictFile', alias: 'd', type: String },
    { name: 'mapFile', alias: 'm', type: String },
    { name: 'keysLine', alias: 'l', type: Number, defaultValue: 0 },
    { name: 'sheetIndex', alias: 'i', type: Number, defaultValue: 0 }
  ])
} catch(e) {
  console.log(usage);
  process.exit(1);
}


/**
 * Save a json file
 *
 * @param {string} filename
 * @param {*} data
 * @returns {void}
 */
function writeJsonFile(filename, data, encoding = 'utf-8', override = false, i = 0) {
  if (i > 0) {
    const _filename = filename.split('.')
    _filename[0] = `${_filename[0]}-${i}`
    filename = _filename.join('.')
  }
  if (!override && existsSync(filename)) {
    const prompt = new Confirm(`O arquivo ${filename} já existe. Deseja sobrescreve-lo?`);
    prompt.run()
      .then(answer => answer ? writeJsonFile(filename, data, encoding, answer) : writeJsonFile(filename, data, encoding, answer, ++i));
    return;
  }
  writeFileSync(path.join(process.cwd(), filename), JSON.stringify(data, null, 2), { encoding });
  console.log(`"${filename}" file saved successfully in "${process.cwd()}".`);
}

/**
 * Open json file and parse
 *
 * @param {string} filename
 * @param {string} [encoding]
 * @returns {*}
 */
function openJson(filename, encoding) {
  try {
    const file = readFileSync(path.join(process.cwd(), filename), { encoding } );
    return JSON.parse(file);
  } catch (e) {
    console.error(`File "${path.join(process.cwd(), filename)}" not found!`);
    process.exit(1);
  }
}

/**
 * Excel file to json with keys translation
 *
 * @param {*} dataFile
 * @param {{[key: string]: string}} excelKeys
 * @param {{[key: string]: string}} [dict]
 * @returns {*}
 */
function toJson(dataFile, excelKeys, dict) {
  let _dataFile = [];
  const translate = !!dict && typeof dict === 'object'
  dataFile.forEach(line => {
    let value = {}
    line.forEach((col, i) => {
      value[translate ? dict[excelKeys[i]] : excelKeys[i]] = col
    });
    _dataFile.push(value);
  });
  return _dataFile;
}

/**
 * Map object keys format
 *
 * @param {any[]} arr
 * @param {{[key: string]: string}} excelKeys
 * @param {{[key: string]: string}} mapKeys
 * @returns
 */
function mapObj(arr, excelKeys, mapKeys) {
  return arr.map(oldValue => {
    let newValue = {};
    excelKeys.forEach(key => {
      let mapedKeys = (mapKeys[key] || "").split('.');
      if (mapedKeys.length === 1) {
        newValue[mapedKeys[0]] = oldValue[key];
        return;
      }
      const olds = [];
      for (let i = 0; i < mapedKeys.length; i++) {
        let isLast = mapedKeys.length - 1 === i;
        let isFirst = i === 0;
        if (isFirst) {
          newValue[mapedKeys[i]] = newValue[mapedKeys[i]] || {};
          olds.push(newValue[mapedKeys[i]]);
        } else {
          olds[olds.length - 1][mapedKeys[i]] = isLast ? oldValue[key] : olds[olds.length - 1][mapedKeys[i]] || {};
          olds.push(olds[olds.length - 1][mapedKeys[i]])
        }
      }
    })
    return newValue;
  });
}

/**
 * Create json object for use in mapFile option
 *
 * @param {string[]} excelKeys
 * @returns
 */
function createMap(excelKeys) {
  let _toMapJson = {};
  excelKeys.forEach(key => _toMapJson[key] = "");
  return _toMapJson;
}

function main(options) {

  const { file: xlsFileName, dictFile: dictFileName, mapFile: mapFileName, keysLine, sheetIndex } = options

  if (!options.file) {
    console.log(usage)
    return 1;
  }

  const fullPath = path.join(process.cwd(), xlsFileName)
  if (!existsSync(fullPath)) {
    console.error(`File "${fullPath}" not found!`);
    return 1
  }

  const file = xlsx.parse(fullPath)[sheetIndex];

  if (!file) {
    console.error(`Sheet index ${sheetIndex} not found.`);
    return 1;
  }

  const dictKeys = dictFileName ? openJson(dictFileName) : undefined;
  const mapKeys = mapFileName ? openJson(mapFileName) : undefined;
  const excelKeys = file.data[keysLine];
  const data = file.data.slice(keysLine+1);
  // Transform excel to JSON
  let outputJson = toJson(data, excelKeys, dictKeys);
  // Map keys
  if (!!options.mapFile) {
    outputJson = mapObj(outputJson, excelKeys, mapKeys)
  }
  // Save files
  writeJsonFile('output.json', outputJson);
  if (options.createMap) {
    writeJsonFile('map_keys.json', createMap(excelKeys));
  }
  if (options.createDict) {
    writeJsonFile('dict.json', createMap(excelKeys));
  }
  return 0;

}

if (options.help) {
  console.log(usage);
  process.exit(0);
}

main(options)
