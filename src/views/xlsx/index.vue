<template>
  <div class="xlsx">
    <input class="file_input"
      type="file" ref="fileRef" @change="onfile($event)"
      accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/vnd.ms-excel"
    >
    <button @click="selectFile()">导入</button>
    <button>导出</button>

    <component v-bind:is="sumLogs"></component>

    <div class="show_sheet">
      <ul>
        <li v-for="(val, i) in sheetNames" :key="i">
          <button class="change" @click="this.changeContent(i)" >{{val}}</button>
        </li>
      </ul>
      <table>
        <tr class="row">
          <th v-for="(val, k) in content[0]" :key="k">
            {{ k }}
          </th>
        </tr>
        <tr v-for="(item, i) in content" :key="i">
          <td v-for="(val, k) in item" :key="k">
            {{ val }}
          </td>
        </tr>
      </table>
    </div>
  </div>
</template>

<script>
import XLSX from 'xlsx'
import sumLogs from '../sumLogs/'

export default {
  name: 'xlsx',
  el: '#app',
  components: {
    sumLogs
  },
  data() {
    return {
      option: {
        n: null,
        r: null,
        o: null,
        fn: null,
      },
      project_exec_list: [

      ],
      sheetNames: [],
      sheets: [],
      content: [],
      file_name: '',
      disable_save: false,
      fileRef: null
    }
  },
  mounted() {
    this.fileRef = this.$refs['fileRef']
    // document.getElementById('export').addEventListener('click', () => {
    //   this.exportExcel('e');
    // })
  },
  methods: {
    changeContent: function (index = 0) {
      const worksheet = this.sheets[this.sheetNames[index]];
      this.content = XLSX.utils.sheet_to_json(worksheet);
    },
    selectFile: function () {
      this.fileRef.click();
    },
    onfile: async function (e) {
      const files = e.target.files; if (files.length === 0) return;
      const f = files[0];
      let re, le;
      eval("re = /xlsx/g; le = /xls/g;");
      if (!re.test(f.name) && !le.test(f.name)) {
        alert(`仅支持读取xlsx或xls格式！`);
        return;
      }
      const workbook = await this.readLocalFile(f);
      this.content = await this.readWorkbook(workbook);

      this.file_name = f.name;
      await this.changeContent();
      console.log(this.content);
      alert('操作成功');
    },
    // 读取本地excel文件
    readLocalFile: function (file) {
      return new Promise((resolve) => {
        var reader = new FileReader();
        reader.onload = async function (e) {
          var data = e.target.result;
          var workbook = await XLSX.read(data, { type: 'binary' });

          resolve(workbook);
        }
        reader.readAsBinaryString(file);
      })
    },
    readWorkbook: function (workbook) {
      this.sheetNames = workbook.SheetNames; // 工作表名称集合
      this.sheets = workbook.Sheets
    },
    openDownloadDialog: function (url, saveName) {
      if (typeof url == 'object' && url instanceof Blob) {
        url = URL.createObjectURL(url); // 创建blob地址
      }
      var aLink = document.createElement('a');
      aLink.href = url;
      aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
      var event;
      if (window.MouseEvent) event = new MouseEvent('click');
      else {
        event = document.createEvent('MouseEvents');
        event.initMouseEvent('click', true, false, window, 0, 0, 0, 0, 0, false, false, false, false, 0, null);
      }
      aLink.dispatchEvent(event);
    },
    // 将一个sheet转成最终的excel文件的blob对象，然后利用URL.createObjectURL下载
    sheet2blob(sheet, sheetName) {
      sheetName = sheetName || 'sheet1';
      var workbook = {
        SheetNames: [sheetName],
        Sheets: {}
      };
      workbook.Sheets[sheetName] = sheet;
      // 生成excel的配置项
      var wopts = {
        bookType: 'xlsx', // 要生成的文件类型
        bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
        type: 'binary',
        password: 'test',
      };
      var wbout = XLSX.write(workbook, wopts);
      var blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
      // 字符串转ArrayBuffer
      function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
      }
      return blob;
    },
    exportExcel: function (option) {
      if (this.content.length === 0) {
        alert("先导入Excel");
        return;
      }
      var temp, blob;

      if (option && option === 'e') {
        blob = new Blob([this.content]);
      } else {
        if (this.option.n === "d") {
            temp = this.content;
            blob = this.sheet2blob(XLSX.utils.json_to_sheet(temp));
        } else {
            blob = new Blob([this.content]);
        }
      }

      const _str = this.file_name.replace(this.option.r, '');
      this.openDownloadDialog(blob, `(${_str})${this.option.o}`);
      this.disable_save = false;
      this.content = null;
    },
    formatDate: function (numb, format) {
      if (numb != undefined) {
        let time = new Date((numb - 1) * 24 * 3600000 + 1)
        time.setYear(time.getFullYear() - 70)
        let year = time.getFullYear() + ''
        let month = time.getMonth() + 1 + ''
        let date = time.getDate() + ''
        if (format && format.length === 1) {
          return year + format + month + format + date
        }
        return year + (month < 10 ? '0' + month : month) + (date < 10 ? '0' + date : date)
      } else {
        return undefined;
      }
    }
  }
}
</script>

<style lang="scss">
  .xlsx {
    .file_input {
      display: none;
    }
  }
</style>