<template>
<div style='margin:24px 0px;'>
<div v-if='files.length > 0'>
  <v-divider style='margin: 24px 0 24px 0;'></v-divider>
  <v-app-bar
    color="accent-4"
    dense dark>
    <v-toolbar-title style='font-size: 0.875rem;'>헤더 지정하기 (*기본값은 첫 행)</v-toolbar-title>
    <v-btn @click='onFixHeader()'>고정</v-btn>
  </v-app-bar>
  <v-simple-table height="200px" v-if='excelViewer != null'>
    <template v-slot:default>
      <tbody>
        <tr v-for="(item, index) in excelViewer" :key="index">
          <td>
          <input ref='check' type='checkbox' style='margin-left:16px;'
          v-model="item['check']" :disabled='fixHeader'>
          </td>
          <td v-for="(value, index) in Object.values(item).splice(0,3)" :key='value+index'>
            {{value}}
          </td>
        </tr>
      </tbody>
    </template>
  </v-simple-table>
  <v-divider style='margin: 24px 0 24px 0;'></v-divider>
  </div>
  </div>
</template>

<script>
import eventBus from './eventBus'
const XLSX = require('xlsx')

export default {
    name : 'viewer',

    created() {
      eventBus.$on('remove-file', e => {
        this.fixHeader = false,
        this.files = e
        this.excelViewer = null
      })

      eventBus.$on('upload-file', async e => {
        this.files = e
        this.excelViewer = await this.onExcelViewer()
      })
    },

    data () {
      return {
        files : [],
        excelViewer : null,
        fixHeader : false,
      }
    },

    methods : {
      async onExcelViewer() {
        if(this.files.length != 0) {
          let result = await this.axios.post('http://localhost:8082/read', {
            path : this.files[0].path,
          })

          let ws = result.data.ws

          let jsonData = XLSX.utils.sheet_to_json(ws).splice(0, 4)
          if(jsonData.length > 0) {
            jsonData.unshift(Object.keys(jsonData[0]))
          }
          return jsonData

        } else return null
      },

      onFixHeader() {
        if(this.fixHeader == false) {
          let result = this.$refs.check.filter( (element) => {
            return element.checked == true
          })
          
          eventBus.$emit('header-number', result.length)

          this.fixHeader = true
        } else if(this.fixHeader == true){
          this.fixHeader = false
        }
      }
    } 
}
</script>
<style>
.v-toolbar__content {
  display:flex!important;
  justify-content:space-between!important;
}
</style>>
