<template>
  <v-card>
    <v-tabs
      v-model="tab"
      background-color="accent-4"
      dark
      @change='resetFiles()'
    >
      <v-tab
        v-for="item in items"
        :key="item.tab"
      >
        {{ item.tab }}
      </v-tab>
    </v-tabs>

    <v-tabs-items v-model="tab">
      <v-tab-item :key='"파일"'>
        <v-card flat style='margin-top: 16px;'>
          <v-card-text id='folder' style='display:flex;align-items: center; justify-content:space-between'>
            <input type='file' accept='.xlsx,.xls'
            @change='getFiles' id='input-files' style='display:none;' ref='files' multiple/>
            
            <h5 style='margin: 6px; max-width: 160px; text-overflow: ellipsis; white-space: nowrap; overflow: hidden;' v-if='files.length > 0'>{{files[0].name}} /</h5>
            <h5 style='margin:6px;'> 총 {{files.length}}개 업로드</h5>
            
            <div>
              <v-btn depressed small color="primary" @click='uploadFiles' style='margin-right: 12px;'>업로드</v-btn>
              <v-btn depressed small color="error" @click='resetFiles()'>취소</v-btn>
            </div>
          </v-card-text>
        </v-card>
      </v-tab-item>

      <v-tab-item :key='"폴더"'>
        <v-card flat style='margin-top: 16px;'>
          <v-card-text id='folder' style='display:flex;align-items: center; justify-content:space-between'>
            <input type='file' @change='getFolder' id='input-folder' style='display:none;' ref='folder' webkitdirectory mozdirectory msdirectory odirectory directory multiple/>
            
            <h5 style='margin: 6px; max-width: 160px; text-overflow: ellipsis; white-space: nowrap; overflow: hidden;' v-if='files.length > 0'>{{files[0].name}} /</h5>
            <h5 style='margin:6px;'> 총 {{files.length}}개 업로드</h5>
            
            <div>
              <v-btn depressed small color="primary" @click='uploadFolder()' style='margin-right: 12px;'>업로드</v-btn>
              <v-btn depressed small color="error" @click='resetFiles()'>취소</v-btn>
            </div>
          </v-card-text>
        </v-card>
      </v-tab-item>
    </v-tabs-items>
  </v-card>
</template>

<script>
import eventBus from './eventBus.js'

export default {
    name : 'file',
    
    data: () => ({
        files : [],
        tab: null,
        items: [
          { tab: '파일' },
          { tab: '폴더' },
        ],
    }),

    methods : {
        getFiles(e) {
          console.log('files-uploaded')
          for(let i = 0; i < e.target.files.length; i++) {
            if ((e.target.files[i].name).match(/\.[a-z]*$/i)[0] == '.xlsx') {
              this.files.push(e.target.files[i])
            } else if ((e.target.files[i].name).match(/\.[a-z]*$/i)[0] == '.xls') {
              this.files.push(e.target.files[i])
            }
          }
          eventBus.$emit('upload-file', this.files)
        },

        getFolder(e) {
          console.log('folder-uploaded')
          for(let i = 0; i < e.target.files.length; i++) {
            if ((e.target.files[i].name).match(/\.[a-z]*$/i)[0] == '.xlsx') {
              this.files.push(e.target.files[i])
            } else if ((e.target.files[i].name).match(/\.[a-z]*$/i)[0] == '.xls') {
              this.files.push(e.target.files[i])
            }
          }
          eventBus.$emit('upload-file', this.files)
        },

        uploadFolder() {
          this.files = []
          this.$refs.folder.click()

          if(this.$refs.folder == undefined) return
          else if(this.$refs.files == undefined) return
          this.$refs.folder.value = ''
          this.$refs.files.value = ''
        },

        uploadFiles() {
          this.files = []
          this.$refs.files.click()

          if(this.$refs.folder == undefined) return
          else if(this.$refs.files == undefined) return
          this.$refs.folder.value = ''
          this.$refs.files.value = ''
        },

        resetFiles() {
          if(this.$refs.folder != undefined) {
            this.files = []
            this.$refs.folder.value = ''
            eventBus.$emit('remove-file', [])

          } else if(this.$refs.files != undefined) {
            this.files = []
            this.$refs.files.value = ''
            eventBus.$emit('remove-file', [])
          }
        },
    }

}
</script>

<style>

</style>