<script setup>
import { computed, ref } from 'vue'
import * as XLSX from 'xlsx'
import {
  NButton,
  NCard,
  NDataTable,
  NEmpty,
  NForm,
  NFormItem,
  NInput,
  NSpace,
  NStatistic,
  NTag,
  NUpload,
} from 'naive-ui'

const COLUMN_INDEX = {
  B: 1,
  C: 2,
  E: 4,
  F: 5,
  H: 7,
  J: 9,
  K: 10,
  N: 13,
  O: 14,
  P: 15,
  Q: 16,
  R: 17,
  S: 18,
}

const tableColumns = [
  { title: '地区号', key: 'regionCode', minWidth: 90 },
  { title: '网点号', key: 'branchCode', minWidth: 90 },
  { title: '网点号名称', key: 'branchName', minWidth: 220 },
  { title: '申报号码', key: 'declarationNo', minWidth: 220 },
  { title: '报送日期', key: 'reportDate', minWidth: 120 },
  { title: '文件种类', key: 'fileType', minWidth: 160 },
  { title: '币种', key: 'currency', minWidth: 130 },
  { title: '金额', key: 'amount', minWidth: 120 },
  { title: '折美元', key: 'usdAmount', minWidth: 120 },
  { title: '客户类型', key: 'customerType', minWidth: 160 },
  { title: '客户名称', key: 'customerName', minWidth: 180 },
  { title: '对方名称', key: 'counterpartyName', minWidth: 180 },
]

const sourceData = ref([])
const keyword = ref('')
const loadedFileName = ref('')
const loading = ref(false)
const processStats = ref({
  totalRows: 0,
  removedByDate: 0,
  removedByCorporate: 0,
  removedByDuplicate: 0,
  keptRows: 0,
  latestOperationDate: '',
  cutoffDate: '',
})
const errorMessage = ref('')

const displayedData = computed(() => {
  const key = keyword.value.trim().toLowerCase()
  if (!key) return sourceData.value

  return sourceData.value.filter((item) => {
    const source = `${item.customerName} ${item.counterpartyName}`.toLowerCase()
    return source.includes(key)
  })
})

const exportDisabled = computed(() => displayedData.value.length === 0)

function normalizeCell(value) {
  if (value === undefined || value === null) return ''
  return String(value).trim()
}

function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate())
  }

  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value)
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d)
    }
  }

  if (typeof value === 'string') {
    const text = value.trim()
    if (!text) return null
    const normalized = text.replace(/[./]/g, '-')
    const parsed = new Date(normalized)
    if (!Number.isNaN(parsed.getTime())) {
      return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate())
    }
  }

  return null
}

function formatDate(date) {
  if (!date) return ''
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  return `${year}-${month}-${day}`
}

function formatDateTime(date) {
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  const hour = String(date.getHours()).padStart(2, '0')
  const minute = String(date.getMinutes()).padStart(2, '0')
  const second = String(date.getSeconds()).padStart(2, '0')
  return `${year}${month}${day}_${hour}${minute}${second}`
}

function mapRowToViewModel(row, index) {
  const reportDate = parseExcelDate(row[COLUMN_INDEX.J])
  return {
    id: `${normalizeCell(row[COLUMN_INDEX.H]) || 'no_declaration'}_${index}`,
    regionCode: normalizeCell(row[COLUMN_INDEX.B]),
    branchCode: normalizeCell(row[COLUMN_INDEX.C]),
    branchName: normalizeCell(row[COLUMN_INDEX.E]),
    declarationNo: normalizeCell(row[COLUMN_INDEX.H]),
    reportDate: formatDate(reportDate) || normalizeCell(row[COLUMN_INDEX.J]),
    fileType: normalizeCell(row[COLUMN_INDEX.K]),
    currency: normalizeCell(row[COLUMN_INDEX.N]),
    amount: normalizeCell(row[COLUMN_INDEX.O]),
    usdAmount: normalizeCell(row[COLUMN_INDEX.P]),
    customerType: normalizeCell(row[COLUMN_INDEX.Q]),
    customerName: normalizeCell(row[COLUMN_INDEX.R]),
    counterpartyName: normalizeCell(row[COLUMN_INDEX.S]),
  }
}

async function parseWorkbook(file) {
  const buffer = await file.arrayBuffer()
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true })
  if (!workbook.SheetNames.length) {
    throw new Error('文件中没有可用工作表')
  }

  const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
  const rows = XLSX.utils.sheet_to_json(firstSheet, {
    header: 1,
    defval: '',
  })

  const bodyRows = rows.slice(1).filter((row) => row.some((cell) => normalizeCell(cell)))
  if (!bodyRows.length) {
    throw new Error('文件没有可处理的数据行')
  }

  const operationDates = bodyRows
    .map((row) => parseExcelDate(row[COLUMN_INDEX.F]))
    .filter((date) => Boolean(date))

  if (!operationDates.length) {
    throw new Error('未检测到业务办理日期（F列）')
  }

  const latestOperationDate = new Date(
    Math.max(...operationDates.map((date) => date.getTime()))
  )
  const cutoffDate = new Date(latestOperationDate)
  cutoffDate.setFullYear(cutoffDate.getFullYear() - 1)

  const seenDeclarationNos = new Set()
  const filteredRows = []
  let removedByDate = 0
  let removedByCorporate = 0
  let removedByDuplicate = 0

  bodyRows.forEach((row, index) => {
    const operationDate = parseExcelDate(row[COLUMN_INDEX.F])
    if (!operationDate || operationDate < cutoffDate) {
      removedByDate += 1
      return
    }

    const customerType = normalizeCell(row[COLUMN_INDEX.Q])
    if (customerType.includes('对公用户')) {
      removedByCorporate += 1
      return
    }

    const declarationNo = normalizeCell(row[COLUMN_INDEX.H])
    if (declarationNo && seenDeclarationNos.has(declarationNo)) {
      removedByDuplicate += 1
      return
    }

    if (declarationNo) {
      seenDeclarationNos.add(declarationNo)
    }

    filteredRows.push(mapRowToViewModel(row, index))
  })

  return {
    filteredRows,
    stats: {
      totalRows: bodyRows.length,
      removedByDate,
      removedByCorporate,
      removedByDuplicate,
      keptRows: filteredRows.length,
      latestOperationDate: formatDate(latestOperationDate),
      cutoffDate: formatDate(cutoffDate),
    },
  }
}

async function handleBeforeUpload({ file }) {
  const rawFile = file.file
  if (!rawFile) return false

  const lowerName = rawFile.name.toLowerCase()
  if (!lowerName.endsWith('.xlsx') && !lowerName.endsWith('.xls')) {
    errorMessage.value = '仅支持上传 .xlsx / .xls 文件'
    return false
  }

  loading.value = true
  errorMessage.value = ''
  try {
    const { filteredRows, stats } = await parseWorkbook(rawFile)
    sourceData.value = filteredRows
    processStats.value = stats
    loadedFileName.value = rawFile.name
    keyword.value = ''
  } catch (error) {
    sourceData.value = []
    processStats.value = {
      totalRows: 0,
      removedByDate: 0,
      removedByCorporate: 0,
      removedByDuplicate: 0,
      keptRows: 0,
      latestOperationDate: '',
      cutoffDate: '',
    }
    errorMessage.value = error instanceof Error ? error.message : '文件处理失败'
  } finally {
    loading.value = false
  }

  return false
}

function handleSearch() {
  keyword.value = keyword.value.trim()
}

function handleReset() {
  keyword.value = ''
}

function handleExport() {
  if (!displayedData.value.length) return

  const exportRows = displayedData.value.map((item) => ({
    地区号: item.regionCode,
    网点号: item.branchCode,
    网点号名称: item.branchName,
    申报号码: item.declarationNo,
    报送日期: item.reportDate,
    文件种类: item.fileType,
    币种: item.currency,
    金额: item.amount,
    折美元: item.usdAmount,
    客户类型: item.customerType,
    客户名称: item.customerName,
    对方名称: item.counterpartyName,
  }))

  const workbook = XLSX.utils.book_new()
  const worksheet = XLSX.utils.json_to_sheet(exportRows)
  XLSX.utils.book_append_sheet(workbook, worksheet, '处理结果')

  XLSX.writeFile(workbook, `外汇处理结果_${formatDateTime(new Date())}.xlsx`)
}
</script>

<template>
  <div class="page">
    <div class="bg-shape bg-shape-left"></div>
    <div class="bg-shape bg-shape-right"></div>

    <div class="container">
      <header class="hero">
        <h1>歪比巴卜外汇清洗机</h1>
        <p>上传 Excel 后自动清洗，再按收款人/汇款人名称查询与导出当前结果。</p>
      </header>

      <NCard class="card" title="1. 上传与预处理">
        <NSpace vertical :size="16">
          <NUpload
            accept=".xlsx,.xls"
            :default-upload="false"
            :max="1"
            :show-file-list="true"
            @before-upload="handleBeforeUpload">
            <NButton type="primary" :loading="loading">选择并处理表格
            </NButton>
          </NUpload>

          <div class="meta-row">
            <NTag v-if="loadedFileName" type="success" size="small">
              当前文件：{{ loadedFileName }}
            </NTag>
            <NTag v-if="processStats.latestOperationDate" type="info"
              size="small">
              日期窗口：{{ processStats.cutoffDate }} 至
              {{ processStats.latestOperationDate }}
            </NTag>
          </div>

          <p v-if="errorMessage" class="error">{{ errorMessage }}</p>

          <div class="stats-grid">
            <NStatistic label="源数据行数"
              :value="processStats.totalRows" />
            <NStatistic label="剔除(一年外)"
              :value="processStats.removedByDate" />
            <NStatistic label="剔除(对公用户)"
              :value="processStats.removedByCorporate" />
            <NStatistic label="剔除(重复申报号)"
              :value="processStats.removedByDuplicate" />
            <NStatistic label="保留行数" :value="processStats.keptRows" />
          </div>
        </NSpace>
      </NCard>

      <NCard class="card" title="2. 查询与导出">
        <NForm inline label-placement="left">
          <NFormItem label="收款人/汇款人">
            <NInput
              v-model:value="keyword"
              clearable
              placeholder="输入客户名称或对方名称"
              @keydown.enter="handleSearch" />
          </NFormItem>
          <NFormItem>
            <NSpace>
              <NButton type="primary" @click="handleSearch">查询
              </NButton>
              <NButton @click="handleReset">重置</NButton>
              <NButton type="success" :disabled="exportDisabled"
                @click="handleExport">
                导出当前显示数据
              </NButton>
            </NSpace>
          </NFormItem>
        </NForm>
      </NCard>

      <NCard class="card" title="3. 结果表">
        <template v-if="displayedData.length">
          <NDataTable
            :columns="tableColumns"
            :data="displayedData"
            :bordered="false"
            size="small"
            :max-height="560"
            :scroll-x="1800" />
        </template>
        <NEmpty v-else description="暂无数据，请先上传并处理表格" />
      </NCard>
    </div>
  </div>
</template>

<style scoped>
.page {
  min-height: 100vh;
  padding: 28px 14px 36px;
  background:
    radial-gradient(circle at 8% 0%, rgba(13, 148, 136, 0.18), transparent 40%),
    radial-gradient(circle at 92% 100%, rgba(15, 23, 42, 0.14), transparent 36%),
    linear-gradient(165deg, #f6fbff 0%, #f8fafb 52%, #eef7ff 100%);
  position: relative;
  overflow: hidden;
}

.bg-shape {
  position: absolute;
  z-index: 0;
  border-radius: 50%;
  filter: blur(3px);
}

.bg-shape-left {
  width: 220px;
  height: 220px;
  left: -80px;
  top: 120px;
  background: rgba(8, 145, 178, 0.2);
}

.bg-shape-right {
  width: 240px;
  height: 240px;
  right: -90px;
  bottom: 70px;
  background: rgba(3, 105, 161, 0.16);
}

.container {
  max-width: 1360px;
  margin: 0 auto;
  position: relative;
  z-index: 1;
  display: grid;
  gap: 16px;
}

.hero {
  border-left: 5px solid #0f766e;
  padding: 14px 16px;
  border-radius: 12px;
  background: rgba(255, 255, 255, 0.75);
  backdrop-filter: blur(6px);
}

.hero h1 {
  font-size: 28px;
  line-height: 1.2;
  font-weight: 700;
  color: #0f172a;
}

.hero p {
  margin-top: 8px;
  color: #334155;
}

.card {
  border-radius: 14px;
  background: rgba(255, 255, 255, 0.85);
  backdrop-filter: blur(5px);
}

.stats-grid {
  display: grid;
  grid-template-columns: repeat(5, minmax(0, 1fr));
  gap: 10px;
}

.meta-row {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
}

.error {
  color: #dc2626;
  font-size: 13px;
}

@media (max-width: 960px) {
  .stats-grid {
    grid-template-columns: repeat(2, minmax(0, 1fr));
  }

  .hero h1 {
    font-size: 24px;
  }
}

@media (max-width: 640px) {
  .page {
    padding: 16px 10px 24px;
  }

  .stats-grid {
    grid-template-columns: 1fr;
  }
}
</style>
