# excel 前端导出插件

## 安装

### 切换到gxnpm源

```bash
nrm use gxnpm
```

- npm:

 ```shell
 npm install exceljs-plugin
 ```

- yarn:  

```shell
yarn add exceljs-plugin
```

### 配置

修改项目根目录下的`angular.json`文件

查找`architect`配直节，在下面的`scripts`内加入`"node_modules/exceljs/dist/exceljs.min.js"`

添加后的效果

```json
 "scripts": [
    "node_modules/@antv/g2/build/g2.js",
    "node_modules/@antv/data-set/dist/data-set.min.js",
    "node_modules/@antv/g2-plugin-slider/dist/g2-plugin-slider.min.js",
    "node_modules/ajv/dist/ajv.bundle.js",
    "node_modules/qrious/dist/qrious.min.js",
    "node_modules/exceljs/dist/exceljs.min.js"
 ]
```

## 使用

在要导出的代码文件中引用包

```javascript
import { GaoxinExcelExport } from 'exceljs-plugin';
```

> `GaoxinExcelExport`下有两个方法，一个是直接导出dom(table dom),一个是导出数据

- 直接导出DOM

页面表格是什么样的,导出就是什么样的，适用于单页展示的复杂表格，只会展示当前页的表格，对于分页的，只会展示第一页。

`exportByDom(dom: any, fileName: string): void`

参数说明

| 参数     | 说明                                    |
| -------- | --------------------------------------- |
| dom      | html table的dom,在ng下是nativeElement   |
| fileName | 文件名,导出时候会自动加入xlsx作为后缀名 |

- 直接导出DOM

将数据导出成excel,适用于数据多的情况下，代码直接将数据传入即可

`export(option: ExcelPluginOption): void`

参数说明

| 参数   | 说明                              |
| ------ | --------------------------------- |
| option | 一个`ExcelPluginOption`类型的选项 |

`ExcelPluginOption` 说明

| 参数       | 类型       | 说明                                                      | 是否必填 |
| ---------- | ---------- | --------------------------------------------------------- | -------- |
| fileName   | `string`   | 文件名                                                    | 是       |
| data       | `[]`       | 数据,对象数组                                             | 是       |
| header     | `Header[]` | 表头，对象数组，如果不填，默认会使用data里面的key作为表头 | 否       |
| mergeCells | `string[]` | 要合并的单元格,遵从excel的格子名称                        | 否       |

- `header`的类型是`Header`

`Header` 说明

| 参数   | 类型     | 说明                                 | 是否必填 |
| ------ | -------- | ------------------------------------ | -------- |
| header | `string` | 表头，写你想要的名字既可             | 是       |
| key    | `string` | key,对应数据的key                    | 是       |
| width  | `number` | 列宽，不写，默认是15                 | 否       |
| style  | `Style`  | 样式，目前只支持上下对齐，左右对齐等 | 否       |

`Style.alignment` 说明

| 参数       | 类型     | 说明                                    | 是否必填 |
| ---------- | -------- | --------------------------------------- | -------- |
| horizontal | `string` | 左右对齐 `'left' | 'center' | 'right' ` | 否       |
| vertical   | `string` | 上下对齐 `'top' | 'middle' | 'bottom' ` | 否       |

一个带`header`的标准例子

```javascript
const header: Header[] = [
    { header: '标题1', key: 'a' },
    { header: '标题2', key: 'b', style: { alignment: { horizontal: 'center' } } },
  ];
const data: any[] = [
    { a: 'hello', b: 'world' },
    { a: 'nihao', b: '还行' },
  ];  
const  opt={
  fileName: 'test',
  header: header,
  data: data,
};
GaoxinExcelExport.export(opt);  
```

通过head可以控制显示的列数,比如数据里面有10列,`header`有两列，那么只会导出这2列

## 例子😆

下面使用`NG-ZORRO`的`nz-table`组件做一个例子

- ### test.component.html

```html
<div nz-row nzGutter="8">
  <div nz-col nzSpan="24">
  <!-- 导出dom按钮 -->
    <button nz-button (click)="export()">
      <i nz-icon type="export"> </i>
      exportDom
    </button>
    <!-- 导出数据按钮 -->
    <button nz-button (click)="exportData()">
      <i nz-icon type="export"> </i>
      exportData
    </button>
  </div>
</div>
<div nz-row nzGutter="8">
  <div nz-col nzSpan="24">
  <!-- 表格实例 -->
    <nz-table #table [nzData]="listOfData" nzBordered>
      <thead>
        <tr>
          <th colspan="8">党政机关办公用房清理腾退情况统计表</th>
        </tr>
        <tr>
          <th rowspan="2" colspan="2">办公用房类型</th>
          <th colspan="4">基本办公用房（使用面积）</th>
          <th rowspan="2">附属用房<br />（建筑面积）</th>
          <th rowspan="2">备 注</th>
        </tr>
        <tr>
          <th>办公室</th>
          <th>服务用房</th>
          <th>设备用房</th>
          <th>小 计</th>
        </tr>
      </thead>
      <tbody>
        <tr *ngFor="let data of table.data; index as i">
          <td>{{ data.key }}</td>
          <td>{{ data.name }}</td>
          <td>{{ data.age }}</td>
          <td>{{ data.tel }}</td>
          <td>{{ data.phone }}</td>
          <td>{{ data.address }}</td>
          <td>{{ data.name }}</td>
          <td>{{ data.name }}</td>
        </tr>
      </tbody>
    </nz-table>
  </div>
</div>
```

> ### test.component.ts

```javascript
import { Component, OnInit, ViewChild } from '@angular/core';
import { GaoxinExcelExport } from 'exceljs-plugin';
import { NzTableComponent } from 'ng-zorro-antd';
import { ExcelPluginOption } from 'exceljs-plugin/dist/interface';
@Component({
  selector: 'app-test',
  templateUrl: './test.component.html',
  styleUrls: ['./test.component.less'],
})
export class TestComponent implements OnInit {
  listOfData = [
    {
      key: '1',
      name: 'John Brown',
      age: 32,
      tel: '0571-22098909',
      phone: 18889898989,
      address: 'New York No. 1 Lake Park',
    },
    {
      key: '2',
      name: 'Jim Green',
      tel: '0571-22098333',
      phone: 18889898888,
      age: 42,
      address: 'London No. 1 Lake Park',
    },
    {
      key: '3',
      name: 'Joe Black',
      age: 32,
      tel: '0575-22098909',
      phone: 18900010002,
      address: 'Sidney No. 1 Lake Park',
    },
    {
      key: '4',
      name: 'Jim Red',
      age: 18,
      tel: '0575-22098909',
      phone: 18900010002,
      address: 'London No. 2 Lake Park',
    },
    {
      key: '5',
      name: 'Jake White',
      age: 18,
      tel: '0575-22098909',
      phone: 18900010002,
      address: 'Dublin No. 2 Lake Park',
    },
  ];
  constructor() {}
  //ng8的用法,ng7:@ViewChild('table')
  @ViewChild('table', { static: false })
  table: NzTableComponent;
  export() {
    GaoxinExcelExport.exportByDom(this.table.tableMainElement.nativeElement, 'hello');
  }

  exportData() {
    const opt: ExcelPluginOption = {
      filName: 'helloData',
      data: this.listOfData,
      mergeCells: ['D2:D6'], //如果需要合并单元格
    };
    GaoxinExcelExport.export(opt);
  }
  ngOnInit() {}
}
```

## 注意的地方

1. `nz-table` 不是dom

`nz-table`的`ViewChild`的类型不是`ElementRef`,而是`NzTableComponent`

所以要使用

```javaScript
this.table.tableMainElement.nativeElement
```

- 单元格内嵌其他元素

插件无法判断`<td>`内部的dom元素(目前仅可以判断一层的input text);

所以如果存在如下情况，可以使用一个通用的class:`display-excel`来处理

> 例子:

```html
<td>
<!-- 这个隐藏的标签来绑定要导出的数据值 -->
<span class="display-excel" style="dispaly:none">{{value}}</span>
<!-- 其他标签是页面显示的内容，比如按钮或者其他元素 -->
<button>...
</td>
```

## 问题

[bug提交](http://172.72.100.37:13530/SoftwareDevelopment/exceljs-plugin/issues)
