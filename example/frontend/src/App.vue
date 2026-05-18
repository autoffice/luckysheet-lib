<template>
  <div class="app-container">
    <header class="app-header">
      <div class="header-left">
        <span class="app-title">📊 Luckysheet-lib 导入导出示例</span>
      </div>
      <div class="header-actions">
        <label class="btn btn-primary">
          <input type="file" @change="handleFileUpload" accept=".xlsx,.xls" style="display: none" />
          📁 上传 Excel
        </label>
        <button class="btn btn-success" @click="loadSampleData">
          📝 加载示例
        </button>
        <button class="btn btn-export" @click="exportToExcel">
          💾 导出 Excel
        </button>
      </div>
    </header>

    <div id="luckysheet" class="luckysheet-container"></div>

    <transition name="toast-fade">
      <div v-if="statusMessage" :class="['toast', statusType]">
        {{ statusMessage }}
      </div>
    </transition>
  </div>
</template>

<script>
import axios from 'axios'

export default {
  name: 'App',
  data() {
    return {
      statusMessage: '',
      statusType: 'info',
      luckysheetInstance: null
    }
  },
  mounted() {
    this.initLuckysheet()
    this.loadSampleData()
  },
  methods: {
    initLuckysheet() {
      const options = {
        container: 'luckysheet',
        title: '工作簿',
        lang: 'zh',
        showinfobar: false,
        data: [{
          name: 'Sheet1',
          color: '',
          status: 1,
          order: 0,
          data: [],
          config: {},
          index: 0
        }]
      }

      if (window.luckysheet) {
        window.luckysheet.create(options)
      } else {
        this.showStatus('Luckysheet 加载失败，请刷新页面', 'error')
      }
    },

    async handleFileUpload(event) {
      const file = event.target.files[0]
      if (!file) return

      this.showStatus('正在上传文件...', 'info')

      const formData = new FormData()
      formData.append('file', file)

      try {
        const response = await axios.post('/api/excel/upload', formData, {
          headers: {
            'Content-Type': 'multipart/form-data'
          }
        })

        const luckysheetData = response.data
        this.loadDataToLuckysheet(luckysheetData)
        this.showStatus(`文件 "${file.name}" 导入成功`, 'success')

        event.target.value = ''
      } catch (error) {
        console.error('上传失败:', error)
        this.showStatus('上传失败: ' + (error.response?.data?.error || error.message), 'error')
      }
    },

    async loadSampleData() {
      this.showStatus('正在加载示例数据...', 'info')

      try {
        const response = await axios.get('/api/excel/sample')
        const luckysheetData = response.data
        this.loadDataToLuckysheet(luckysheetData)
        this.showStatus('示例数据加载成功', 'success')
      } catch (error) {
        console.error('加载示例数据失败:', error)
        this.showStatus('加载示例数据失败: ' + error.message, 'error')
      }
    },

    loadDataToLuckysheet(data) {
      if (window.luckysheet) {
        window.luckysheet.destroy()

        const options = {
          container: 'luckysheet',
          title: data.info?.name || '工作簿',
          lang: 'zh',
          showinfobar: false,
          data: data.sheets || []
        }

        window.luckysheet.create(options)
      }
    },

    async exportToExcel() {
      this.showStatus('正在导出 Excel...', 'info')

      try {
        const luckysheetData = window.luckysheet.getAllSheets()

        const luckyFile = {
          info: {
            name: '导出文件',
            creator: 'Luckysheet Example',
            createdTime: new Date().toISOString()
          },
          sheets: luckysheetData
        }

        const response = await axios.post('/api/excel/export', luckyFile, {
          responseType: 'blob',
          headers: {
            'Content-Type': 'application/json'
          }
        })

        const blob = new Blob([response.data], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        })
        const url = window.URL.createObjectURL(blob)
        const link = document.createElement('a')
        link.href = url
        link.download = `导出文件_${new Date().getTime()}.xlsx`
        document.body.appendChild(link)
        link.click()
        document.body.removeChild(link)
        window.URL.revokeObjectURL(url)

        this.showStatus('Excel 导出成功', 'success')
      } catch (error) {
        console.error('导出失败:', error)
        this.showStatus('导出失败: ' + error.message, 'error')
      }
    },

    showStatus(message, type = 'info') {
      this.statusMessage = message
      this.statusType = type

      if (type !== 'error') {
        setTimeout(() => {
          this.statusMessage = ''
        }, 2500)
      } else {
        setTimeout(() => {
          this.statusMessage = ''
        }, 5000)
      }
    }
  }
}
</script>

<style>
/* 全局样式: 确保 Luckysheet 容器能正确计算高度 */
html,
body,
#app {
  height: 100%;
  margin: 0;
  padding: 0;
  overflow: hidden;
}
</style>

<style scoped>
.app-container {
  display: flex;
  flex-direction: column;
  height: 100vh;
  width: 100vw;
  background: #f5f5f5;
  overflow: hidden;
}

.app-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 16px;
  height: 40px;
  flex: 0 0 40px;
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color: white;
  box-shadow: 0 1px 3px rgba(0, 0, 0, 0.12);
  z-index: 10;
}

.header-left {
  display: flex;
  align-items: center;
  gap: 12px;
}

.app-title {
  font-size: 14px;
  font-weight: 600;
  letter-spacing: 0.3px;
}

.header-actions {
  display: flex;
  gap: 8px;
  align-items: center;
}

.btn {
  padding: 4px 12px;
  border: none;
  border-radius: 4px;
  font-size: 12px;
  font-weight: 500;
  line-height: 22px;
  cursor: pointer;
  transition: background 0.2s ease, transform 0.1s ease;
  display: inline-flex;
  align-items: center;
}

.btn:hover {
  transform: translateY(-1px);
}

.btn:active {
  transform: translateY(0);
}

.btn-primary {
  background: rgba(255, 255, 255, 0.18);
  color: white;
}

.btn-primary:hover {
  background: rgba(255, 255, 255, 0.28);
}

.btn-success {
  background: #48bb78;
  color: white;
}

.btn-success:hover {
  background: #38a169;
}

.btn-export {
  background: #ed8936;
  color: white;
}

.btn-export:hover {
  background: #dd6b20;
}

.luckysheet-container {
  flex: 1;
  min-height: 0;
  position: relative;
}

/* 浮动提示: 不占布局, 从右上角滑入 */
.toast {
  position: fixed;
  top: 52px;
  right: 16px;
  padding: 8px 16px;
  border-radius: 4px;
  font-size: 13px;
  z-index: 9999;
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
  max-width: 360px;
}

.toast.info {
  background: #ebf8ff;
  color: #2c5282;
  border-left: 3px solid #4299e1;
}

.toast.success {
  background: #f0fff4;
  color: #22543d;
  border-left: 3px solid #48bb78;
}

.toast.error {
  background: #fff5f5;
  color: #742a2a;
  border-left: 3px solid #f56565;
}

.toast-fade-enter-active,
.toast-fade-leave-active {
  transition: opacity 0.25s ease, transform 0.25s ease;
}

.toast-fade-enter-from,
.toast-fade-leave-to {
  opacity: 0;
  transform: translateY(-8px);
}
</style>
