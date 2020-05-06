//typescript实现断点测试，配置在src下才生效
import { ExcelPluginOptionUserDefine } from './interface';
import { Cell, Worksheet, Workbook } from 'exceljs';
const ExcelJS = require('exceljs');
const opt: ExcelPluginOptionUserDefine = {
  fileName: '工业年报表1-8-1：企业能源计量器具配备和管理情况表20200506',
  data: [
    ['工业年报表1-8-1:企业能源计量器具配备和管理情况表'],
    ['公司名称:青海华电大通发电有限公司', '', '', '', '年份:2020'],
    ['等级', '序号', '能源种类及限定值', '计量器具类别', '运行状态', '安装使用地点', '是否在检定周期内'],
    ['进出用能单位', 1, null, '电能表', '良好', null, '否'],
    ['进出用能单位', 2, '洗精煤', null, '良好', null, '否'],
    ['进出用能单位', 3, null, '衡器', '良好', null, '否'],
    ['小计', '', '应配数量（台）', '实配数量（台）', '配备率（%）', '完好率（%）', '检定率（%）'],
    ['', '', 3, 2, 66.67, 100, 44],
    [],
    ['等级', '序号', '能源种类及限定值', '计量器具类别', '运行状态', '安装使用地点', '是否在检定周期内'],
    ['进出主要次级用能单位', 1, null, '衡器', '停用', '厂房1', '否'],
    ['进出主要次级用能单位', 2, null, '衡器', '良好', '', '是'],
    ['进出主要次级用能单位', 3, null, '衡器', '良好', '厂房1111222', '是'],
    ['进出主要次级用能单位', 4, '电力', '电能表', '维护', '未填写', '否'],
    ['进出主要次级用能单位', 5, '天然气', '水流量表(装置)', '维护', '而我却若王二翁人', '是'],
    ['小计', '', '应配数量（台）', '实配数量（台）', '配备率（%）', '完好率（%）', '检定率（%）'],
    ['', '', 10, 5, 50, 12, 78],
    [],
    ['等级', '序号', '能源种类及限定值', '', '应配数', '实配数', '完好数'],
    ['主要用能设备', 1, null, '', 10, 4, 2],
    ['主要用能设备', 2, '洗精煤', '', 11, 11, 7],
    ['主要用能设备', 3, null, '', 422, 112, 112],
    ['小计', '', '应配数量（台）', '', '实配数量（台）', '配备率（%）', '完好率（%）'],
    ['', '', 443, '', 127, 28.67, 95.28],
    [],
    ['项目', '', '要求', '', '', '', '是或否'],
    ['能源计量制度', '', '是否建立能源计量制度，并形成文件', '', '', '', '否'],
    ['能源计量人员', '', '是否有人负责能源器具的管理', '', '', '', '否'],
    ['', '', '是否有专人负责主要次级用能单位和主要用能设备的管理', '', '', '', '否'],
    ['能源计量制度', '', '是否有完整的能源计量器具一览表', '', '', '', '否'],
    ['', '', '是否建立符合规定的能源计量器具档案情况', '', '', '', '否'],
    ['能源计量数据', '', '是否建立能源统计报表制度', '', '', '', '否'],
    ['', '', '是否有用于能源计量数据记录的标准表格样式', '', '', '', '否'],
    ['', '', '是否实现了能源计量数据的网络化管理', '', '', '', '否'],
    [''],
    ['能源管理负责人:李三', '', '填报人:张玲芝', '', '电话:13333333333', '', '填报日期:2020-04-14'],
    ['说明'],
    ['1．次级用能单位：是用能单位下属的能源核算单位。'],
    [
      '2．主要次级用能单位、主要用能设备应按照GB17167-2006《用能单位能源计量器具配备和管理通则》中有关主要次级用能单位、主要用能设备能耗（或功率）限定值进行判定。',
    ],
    ['3．计量器具类别：衡器、电能表、油流量表（装置）、气体流量表（装置）、水流量表（装置）等；（采取选项形式）'],
    ['4．运行状态：良好、维护、停用。'],
    [
      '5．能源种类：指电、煤炭、原油、天然气、焦炭、煤气、热力、成品油、液化石油气、生物质能和其他直接或者通过加工、转换而取得有用能的各种能源。',
    ],
    ['6．填报单位应根据实际情况详细注明计量器具安装使用地点。'],
    ['7．能源计量器具的管理要求依据GB17167-2006《用能单位能源计量器具配备和管理通则》的要求。'],
  ],
  mergeCells: [
    'A1:G1',
    'A2:D2',
    'E2:G2',
    'A7:B8',
    'A9:G9',
    'A16:B17',
    'A18:G18',
    'C19:D19',
    'C20:D20',
    'C21:D21',
    'C22:D22',
    'C23:D23',
    'C24:D24',
    'A23:B24',
    'A25:G25',
    'C26:F26',
    'C27:F27',
    'C28:F28',
    'C29:F29',
    'C30:F30',
    'C31:F31',
    'C32:F32',
    'C33:F33',
    'C34:F34',
    'A26:B26',
    'A27:B27',
    'A28:B29',
    'A30:B31',
    'A32:B34',
    'A35:G35',
    'A37:G37',
    'A38:G38',
    'A39:G39',
    'A40:G40',
    'A41:G41',
    'A42:G42',
    'A43:G43',
    'A44:G44',
  ],
  columnWidth: {
    A: 30,
    B: 30,
    C: 30,
    D: 30,
    E: 30,
    F: 30,
    G: 30,
  },
  columnStyle: {
    A: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
    B: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
    C: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
    D: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
    E: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
    F: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
    G: {
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
      },
    },
  },
  styleToCellList: [
    {
      cellStyle: {
        alignment: {
          horizontal: 'left',
          vertical: 'middle',
        },
        border: null,
      },
      cellList: ['A37', 'A38', 'A39', 'A40', 'A41', 'A42', 'A43', 'A44'],
    },
    {
      cellStyle: {
        border: null,
      },
      cellList: ['A1', 'A2', 'E2'],
    },
  ],
};
// const handel = new HandleForData();
// const header: Header[] = [
//   { header: 'a', key: 'a' },
//   { header: 'bbbb', key: 'b', style: { alignment: { horizontal: 'center' } } },
// ];
// const data = [{ a: 'nihsdfffffffffffffffffffffffffffao', b: 1 }, { a: 'helsdafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafaflo', b: 'world' }];
// const workbookMy = new Excel.Workbook();
// const opt: ExcelPluginOption = {
//   fileName: 'eee',
//   header: header,
//   data: data,
// };
// const workboot = handel.processWorkbook(opt, workbookMy);
//  workboot.xlsx.writeFile('./test.xlsx').then(() => {
//   expect('hello').toBe('hello');
// });
const title = [];
const name = '上报数据项报表';
title.push(name);

const body = [
  ['全厂-二次能源-电力-购入已消费量（电力）', '全厂', '', '', '千瓦时', '', '吨标准煤/万千瓦时', '1.2290', '没有采集'],
];
body;

const head = ['指标名称', '指标范围', '时间', '指标值', '计量单位', '折标值', '折标单位', '折标系数', '备注'];
const result = [title, head, ...body];

const opt1: ExcelPluginOptionUserDefine = {
  fileName: name,
  data: result,
  mergeCells: ['A1:B1'],
  columnWidth: {
    A: 40,
  },
  columnStyle: {
    A: {
      alignment: {
        horizontal: 'center',
      },
    },
  },
};

let workbook: Workbook = new ExcelJS.Workbook();
const sheet: Worksheet = workbook.addWorksheet('sheet1');
sheet.addRows(opt.data);
if (opt.enbaleWrapText !== undefined) {
  this.enbaleWrapText = opt.enbaleWrapText;
}
sheet.eachRow((row, rowNumber) => {
  row.eachCell((cell: Cell, colNumber: number) => {
    cell.alignment = {
      vertical: 'middle',
      horizontal: 'left',
      wrapText: this.enbaleWrapText,
    };
    cell.font = { size: 12, family: 1, bold: false };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' },
    };
  });
});
for (let index = 0; index < sheet.columns.length; index++) {
  const col = sheet.columns[index];
  col.width = 16;
}
if (opt.columnWidth) {
  for (const key in opt.columnWidth) {
    sheet.getColumn(key).width = opt.columnWidth[key];
  }
}
if (opt.mergeCells && opt.mergeCells.length > 0) {
  for (const cell of opt.mergeCells) {
    // @ts-ignore
    sheet.mergeCells(cell);
  }
}

if (opt.columnStyle && Object.keys(opt.columnStyle).length > 0) {
  for (const key in opt.columnStyle) {
    for (const styleKey in opt.columnStyle[key]) {
      sheet.getColumn(key)[styleKey] = opt.columnStyle[key][styleKey];
    }
  }
}

//批量更改某种样式下的单元格数组
if (opt.styleToCellList.length > 0) {
  opt.styleToCellList.forEach((item) => {
    const style = item.cellStyle;
    item.cellList.forEach((cellCode) => {
      for (const styleKey in style) {
        sheet.getCell(cellCode)[styleKey] = style[styleKey];
      }
    });
  });
}

if (opt.cellStyle && Object.keys(opt.cellStyle).length > 0) {
  for (const cellCode in opt.cellStyle) {
    for (const styleKey in opt.cellStyle[cellCode]) {
      sheet.getCell(cellCode)[styleKey] = opt.cellStyle[cellCode][styleKey];
    }
  }
}
workbook.xlsx.writeFile('./test.xlsx').then(() => {
  expect('hello').toBe('hello');
});
