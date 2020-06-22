<template>
  <v-app>
    <v-app-bar app dark>
    Excel Updater
    </v-app-bar>
    
    <v-main style='padding:100px 16px 16px 16px'>
      <div style='display:flex;position:relative;justify-content:center;z-index:999;'>
        <v-alert type='info' v-show='alertUpload' style='position:absolute'>
          {{files.length}} 건의 파일이 성공적으로 업로드 되었습니다.
        </v-alert>

        <v-alert type='success' v-show='alertSuccess' style='position:absolute'>
          {{currentUpdate}}/{{files.length}} 진행 중입니다.
        </v-alert>

        <v-alert type='error' v-show='nullError.length != 0' style='position:absolute'>
          {{nullError}}
        </v-alert>
      </div>

      <file/>

      <viewer/>
      
      <v-card>
        <v-tabs v-model="tab" dark>
          <v-tab
            v-for="item in items"
            :key="item.tab"
          >
            {{ item.tab }}
          </v-tab>
        </v-tabs>

        <v-tabs-items v-model="tab">
          <v-tab-item :key='"전체"'>
            <v-card flat style='margin-top: 16px;'>
              <v-card-text>
              <v-text-field
                label="열 번호(알파벳, ex. A/B/C/D..)"
                v-model='findColumn'
              ></v-text-field>
              <v-textarea
                label="변경 값"
                v-model='changeValue1'
              ></v-textarea>
                <v-btn @click='onAllColumnUpdate()'>전체 변환</v-btn>
              </v-card-text>
            </v-card>
            </v-tab-item>
              
          <v-tab-item :key='"하나"'>
            <v-card flat style='margin-top: 16px;'>
              <v-card-text>
              <v-textarea
                label="찾을 값"
                v-model='findValue'
              ></v-textarea>
              <v-textarea
                label="변경 값"
                v-model='changeValue2'
              ></v-textarea>
                <v-btn @click='onSelectUpdate()'>특정 값 변환</v-btn>
              </v-card-text>
            </v-card>
          </v-tab-item>

          <v-tab-item :key='"중복"'>
            <v-card flat style='margin-top: 16px;'>
              <v-card-text style='display;flex;justify-content:center;'>
                <v-btn @click='onUniqueUpdate()'>상품명 중복 제거 실행</v-btn>
              </v-card-text>
            </v-card>
          </v-tab-item>
        </v-tabs-items>
      </v-card>
    </v-main>
  </v-app>
</template>

<script>
import eventBus from './components/eventBus'

import file from './components/file';
import viewer from './components/viewer';

export default {
  name: 'App',

  components: {
    file, viewer
  },

  data: () => ({
    files : [],
    headerNumber : 1,
    tab: null,
    items: [
      { tab: '전체 열 수정'},
      { tab: '하나의 값 수정'},
      { tab: '상품명 중복 제거'},
    ],

    changeValue1 : null,
    changeValue2 : null,
    findColumn : null,
    findValue : null,
    productColumn : null,

    alertUpload : false,
    alertSuccess : false,
    nullError : '',

    currentUpdate : 0,
  }),

  created() {
    eventBus.$on('remove-file', e => {
      this.files = e
    })

    eventBus.$on('upload-file', e => {
      this.files = e

      this.alertUpload = true

      setTimeout( () => {
        this.alertUpload = false
      }, 1000)
    })

    eventBus.$on('header-number', e => {
      this.headerNumber = e
    })
  },

  methods : {

    async onUniqueUpdate() {
      
      if(this.files.length == 0) {
        this.nullError = '수정할 파일을 선택해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return
      }

      this.alertSuccess = true

      for(let x = 0; x < this.files.length; x ++) {
        this.currentUpdate = x

        let result = await this.axios.post('http://localhost:8082/read', {
          path : this.files[x].path,
        })
      

        for(let y = 0; y <= this.alphaToNum(result.data.maxAlpha); y++) {
          if(result.data.ws[this.numToAlpha(y) + '1'].v == '상품명') {
            this.productColumn = this.numToAlpha(y)
          }
        }

        if(this.productColumn != null) {
          for(let z = (this.headerNumber+1); z <= Number(result.data.maxNumber); z++) {
            if(result.data.ws[this.productColumn + z] != undefined) {
              let array = String(result.data.ws[this.productColumn + z].v).split(' ')

              let uniq = array.reduce( (a,b) => {
                if(a.indexOf(b) <0) a.push(b)
                return a;
              }, [])
              
              uniq = uniq.filter(item => item)

              uniq = uniq.join(" ")

              result.data.ws[this.productColumn + z].v = uniq
              result.data.ws[this.productColumn + z].t = 's'
            }
          }
        }
        
        if (result.data.fileType == '.xls') {
          let final = await this.axios.post('http://localhost:8082/writeXLS', {
            path : this.files[x].path,
            ws : result.data.ws
          })
          console.log(final)
        } else if (result.data.fileType == '.xlsx') {
          let final = await this.axios.post('http://localhost:8082/writeXLSX', {
            path : this.files[x].path,
            wb : result.data.wb,
            ws : result.data.ws

          })
          console.log(final)
        }
      }

      setTimeout( () => {
        this.alertSuccess = false
        this.currentUpdate = 0
        this.headerNumber = 1
      }, 1500)
    },

    async onAllColumnUpdate() {

      if(this.files.length == 0) {
        this.nullError = '수정할 파일을 선택해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return
      } else if(this.findColumn == null) {
        this.nullError = '수정할 열을 입력해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return

      } else if(this.changeValue1 == null) {
        this.nullError = '변경할 값을 입력해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return
      }

      this.alertSuccess = true

      for(let x = 0; x < this.files.length; x ++) {
        this.currentUpdate = x
        let result = await this.axios.post('http://localhost:8082/read', {
          path : this.files[x].path,
        })

        for(let y = (this.headerNumber + 1); y <= Number(result.data.maxNumber); y++) {
          if(result.data.ws[this.findColumn + y] != undefined) {
            result.data.ws[this.findColumn + y].v = this.changeValue1
          } else if ( (( result.data.ws[this.findColumn + y] == undefined ) + (result.data.ws[this.findColumn + (Number(y)+1)]) == undefined) == 0 ) {
            result.data.ws[this.findColumn + y] = {};
            result.data.ws[this.findColumn + y].t = 's';
            result.data.ws[this.findColumn + y].v = this.changeValue1
          }
        }
        
        if (result.data.fileType == '.xls') {
          let final = await this.axios.post('http://localhost:8082/writeXLS', {
            path : this.files[x].path,
            ws : result.data.ws
          })
          console.log(final)
        } else if (result.data.fileType == '.xlsx') {
          let final = await this.axios.post('http://localhost:8082/writeXLSX', {
            path : this.files[x].path,
            wb : result.data.wb,
            ws : result.data.ws

          })
          console.log(final)
        }
      }

      setTimeout( () => {
        this.alertSuccess = false
        this.currentUpdate = 0
        this.findColumn = null
        this.changeValue1 = null
        this.headerNumber = 1
      }, 1500)
    },

    async onSelectUpdate() {
      if(this.files.length == 0) {
        this.nullError = '수정할 파일을 선택해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return

      } else if(this.findValue == null) {
        this.nullError = '찾아서 수정할 값을 입력해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return

      } else if(this.changeValue2 == null) {
        this.nullError = '변경할 값을 입력해 주세요.'
        setTimeout(()=>{
          this.nullError = ''
        }, 1500)
        return
      }

      this.alertSuccess = true

      for(let x = 0; x < this.files.length; x ++) {
        this.currentUpdate = x
        let result = await this.axios.post('http://localhost:8082/read', {
          path : this.files[x].path,
        })

        for(let y = 0; y <= Number(this.alphaToNum(result.data.maxAlpha)); y++) {
          for(let z = (this.headerNumber + 1); z <= Number(result.data.maxNumber); z++) {
            if(result.data.ws[this.numToAlpha(y) + z] != undefined) {
              result.data.ws[this.numToAlpha(y) + z].v = String(result.data.ws[this.numToAlpha(y) + z].v).replace(this.findValue, this.changeValue2)
            }
          }
        }
        
        if (result.data.fileType == '.xls') {
          let final = await this.axios.post('http://localhost:8082/writeXLS', {
            path : this.files[x].path,
            ws : result.data.ws
          })
          console.log(final)
        } else if (result.data.fileType == '.xlsx') {
          let final = await this.axios.post('http://localhost:8082/writeXLSX', {
            path : this.files[x].path,
            wb : result.data.wb,
            ws : result.data.ws

          })
          console.log(final)
        }
      }

      setTimeout( () => {
        this.alertSuccess = false
        this.currentUpdate = 0
        this.findValue = null
        this.changeValue2 = null
        this.headerNumber = 1
      }, 1500)
    },

    alphaToNum(alpha) {
      var i = 0,
      num = 0,
      len = alpha.length;

      for (; i < len; i++) {
        num = num * 26 + alpha.charCodeAt(i) - 0x40;
      }

      return num - 1;
    },

    numToAlpha(num) {

      var alpha = '';

      for (; num >= 0; num = parseInt(num / 26, 10) - 1) {
        alpha = String.fromCharCode(num % 26 + 0x41) + alpha;
      }

      return alpha;
    },
  }
};
</script>
