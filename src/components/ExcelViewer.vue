<template>
  <div style="width: 1920px; height: 1080px;">
    <input type="file" @change="handleFileUpload" accept=".xlsx" />
    <div v-if="uploading" class="progress-bar">
      <div class="progress" :style="{ width: `${progress}%` }"></div>
    </div>
    <div v-if="uploadError" class="error-message">{{ uploadError }}</div>
    <div v-if="uploadSuccess" class="success-message">文件上传成功</div>
    <table v-if="tableDataWithMerged.length">
      <thead>
        <tr>
          <th v-for="(header, index) in headers" :key="index" :style="getHeaderStyle(index)">
            {{ header }}
          </th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(row, rowIndex) in tableDataWithMerged" :key="rowIndex">
          <td
            v-for="(cell, cellIndex) in row"
            :key="cellIndex"
            :rowspan="getMergedRowspan(rowIndex, cellIndex)"
            :colspan="cell.colspan || 1"
            :style="getCellStyle(rowIndex, cellIndex, cell.offset)"
            :class="{
              'merged-cell': isMergedCell(rowIndex, cellIndex),
              'text-center': true
            }"
          >
            {{ formatCellValue(cell.value) }}
        </td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script>
import * as XLSX from 'xlsx'
// import XLSXSTYLE from 'xlsx-style'

export default {
  data () {
    return {
      headers: [],
      tableData: [],
      tableDataWithMerged: [],
      mergedCells: [],
      cellStyles: {},
      uploading: false,
      progress: 0,
      uploadError: '',
      uploadSuccess: false,
      worksheet: null,
      columnsWithData: []
    }
  },
  methods: {
    handleFileUpload (event) {
      const file = event.target.files[0]
      if (!file) return
      if (!file.name.endsWith('.xlsx')) {
        this.uploadError = '请选择一个有效的 Excel 文件'
        return
      }

      this.uploading = true
      this.progress = 0
      this.uploadError = ''
      this.uploadSuccess = false

      const reader = new FileReader()

      reader.onprogress = (e) => {
        if (e.lengthComputable) {
          this.progress = Math.round((e.loaded / e.total) * 100)
        }
      }

      reader.onload = (e) => {
        const data = new Uint8Array(e.target.result)
        try {
          this.parseExcel(data)
          this.uploadSuccess = true
        } catch (error) {
          console.error('文件解析失败:', error)
          this.uploadError = '文件解析失败，请检查文件格式'
        } finally {
          this.uploading = false
        }
      }
      reader.readAsArrayBuffer(file)
    },
    parseExcel (data) {
      console.log('开始解析 Excel 文件')
      const workbook = XLSX.read(data, { type: 'array', cellStyles: true })
      console.log('Workbook:', workbook)
      const firstSheetName = workbook.SheetNames[0]
      this.worksheet = workbook.Sheets[firstSheetName]
      console.log('Worksheet:', this.worksheet)
      const range = XLSX.utils.decode_range(this.worksheet['!ref'])
      console.log('Range:', range)

      const maxRowIndex = range.e.r
      const maxColIndex = range.e.c

      this.extractHeaders(this.worksheet, range, maxColIndex)
      this.extractMergedCells(this.worksheet, maxRowIndex, maxColIndex)
      this.extractCellStyles(this.worksheet, maxRowIndex, maxColIndex)
      this.extractTableData(this.worksheet, range, maxRowIndex, maxColIndex)
      this.processMergedCells()
      this.adjustTableColumns()
    },
    extractHeaders (worksheet, range, maxColIndex) {
      this.headers = []
      for (let colIndex = range.s.c; colIndex <= maxColIndex; colIndex++) {
        const cellAddress = XLSX.utils.encode_cell({ r: 0, c: colIndex })
        const cell = worksheet[cellAddress]
        this.headers.push(cell ? cell.v : '')
      }
    },
    extractMergedCells (worksheet, maxRowIndex, maxColIndex) {
      this.mergedCells = (worksheet['!merges'] || []).map(merge => ({
        startRow: merge.s.r - 1,
        startCol: merge.s.c,
        endRow: merge.e.r - 1,
        endCol: merge.e.c
      }))
      console.log('Merged Cells:', this.mergedCells)
    },
    extractCellStyles (worksheet, maxRowIndex, maxColIndex) {
      this.cellStyles = {}
      for (let rowIndex = 0; rowIndex <= maxRowIndex; rowIndex++) {
        for (let colIndex = 0; colIndex <= maxColIndex; colIndex++) {
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
          const cell = worksheet[cellAddress]
          if (cell && cell.s) {
            const style = cell.s
            // 添加单元格宽度和高度
            const colWidth = worksheet[XLSX.utils.encode_col(colIndex)]?.wch
            const rowHeight = worksheet[XLSX.utils.encode_row(rowIndex)]?.hpt
            style.colWidth = colWidth || 100 // 默认宽度
            style.rowHeight = rowHeight || 20 // 默认高度
            this.cellStyles[`${rowIndex},${colIndex}`] = style
            console.log('Cell Styles:', this.cellStyles)
          }
        }
      }
    },
    extractTableData (worksheet, range, maxRowIndex, maxColIndex) {
      this.tableData = []
      for (let rowIndex = 1; rowIndex <= maxRowIndex; rowIndex++) {
        const row = []
        for (let colIndex = range.s.c; colIndex <= maxColIndex; colIndex++) {
          const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
          const cell = worksheet[cellAddress]
          row.push(cell ? cell.w : '') // 使用 cell.w 来获取格式化后的值
        }
        this.tableData.push(row)
      }
    },
    processMergedCells () {
      this.tableDataWithMerged = this.tableData.map((row, rowIndex) =>
        row.map((cell, cellIndex) => ({
          value: cell,
          merged: this.isMergedCell(rowIndex, cellIndex) && rowIndex === this.findMergeStartRow(rowIndex, cellIndex),
          colspan: this.getMergedColspan(rowIndex, cellIndex)
        }))
      )
      this.fillMergedCellPlaceholders()
    },
    isMergedCell (rowIndex, colIndex) {
      return this.mergedCells.some(merge =>
        merge.startRow <= rowIndex && rowIndex <= merge.endRow &&
        merge.startCol <= colIndex && colIndex <= merge.endCol
      )
    },
    findMergeStartRow (rowIndex, colIndex) {
      return this.mergedCells.find(merge =>
        merge.startRow <= rowIndex && rowIndex <= merge.endRow &&
        merge.startCol === colIndex
      )?.startRow
    },
    getMergedRowspan (rowIndex, colIndex) {
      const merge = this.mergedCells.find(m => m.startRow === rowIndex && m.startCol === colIndex)
      return merge ? (merge.endRow - merge.startRow + 1) : 1
    },
    getMergedColspan (rowIndex, colIndex) {
      const merge = this.mergedCells.find(m => m.startRow === rowIndex && m.startCol === colIndex)
      return merge ? (merge.endCol - merge.startCol + 1) : 1
    },
    // 样式
    getCellStyle (rowIndex, colIndex, offset = 0) {
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex })
      const cell = this.worksheet[cellAddress]
      const style = this.cellStyles[`${rowIndex},${colIndex}`] || {}
      console.log('查看rowIndex:', { rowIndex })
      console.log({ colIndex })
      console.log('查看cellStyles :', this.cellStyles)
      console.log('Cell 查看Style', style)
      console.log('Cell 查看Stylefont', style.font)

      const styles = []

      // // 处理对齐方式
      // if (style.alignment) {
      //   if (style.alignment.horizontal) {
      //     styles.push(`text-align: ${style.alignment.horizontal} !important;`)
      //   }
      //   if (style.alignment.vertical) {
      //     styles.push(`vertical-align: ${style.alignment.vertical} !important;`)
      //   }
      // }

      // 处理背景颜色
      if (style && style.fgColor && style.fgColor.rgb) {
        styles.push(`background-color: #${style.fgColor.rgb} !important;`)
      }

      // 处理字体颜色
      if (style.font && style.font.color && style.font.color.rgb) {
        styles.push(`color: #${style.font.color.rgb} !important;`)
      }

      // 处理字体样式
      if (style.font) {
        if (style.font.bold) {
          styles.push('font-weight: bold !important;')
        }
        if (style.font.italic) {
          styles.push('font-style: italic !important;')
        }
        if (style.font.underline) {
          styles.push('text-decoration: underline !important;')
        }
      }

      // 处理边框
      if (style.border) {
        const handleBorder = (position, styleProp) => {
          if (style.border[position] && style.border[position].style) {
            styles.push(`border-${position}: ${this.getBorderStyle(style.border[position].style)} !important;`)
          }
        }
        handleBorder('top', style.border.top)
        handleBorder('right', style.border.right)
        handleBorder('bottom', style.border.bottom)
        handleBorder('left', style.border.left)
      }

      // 处理单元格宽度和高度
      if (style.colWidth) {
        styles.push(`width: ${style.colWidth}px !important;`)
      }
      if (style.rowHeight) {
        styles.push(`height: ${style.rowHeight}px !important;`)
      }

      // 应用偏移量
      if (offset > 0) {
        const cellWidth = 100 // 假设每个单元格的宽度为100px，可以根据实际情况调整
        styles.push(`left: -${offset * cellWidth}px !important;`)
      }

      // 确保生成的样式字符串格式正确
      const finalStyle = styles.join('').replace(/;/g, '; ').trim() + ';'
      // console.log(`Generated style for cell (${rowIndex}, ${colIndex}):`, finalStyle)
      return finalStyle
    },
    forceRerender () {
      this.$forceUpdate()
    },
    fillMergedCellPlaceholders () {
      const processedCells = Array.from({ length: this.tableData.length }, () => Array(this.tableData[0].length).fill(false))
      const cellOffsets = Array.from({ length: this.tableData.length }, () => Array(this.tableData[0].length).fill(0))

      this.tableDataWithMerged.forEach((row, rowIndex) => {
        row.forEach((cell, cellIndex) => {
          if (cell.merged && rowIndex === this.findMergeStartRow(rowIndex, cellIndex)) {
            const merge = this.mergedCells.find(m => m.startRow === rowIndex && m.startCol === cellIndex)
            const startValue = this.tableData[merge.startRow][merge.startCol]
            for (let r = merge.startRow; r <= merge.endRow; r++) {
              for (let c = merge.startCol; c <= merge.endCol; c++) {
                if (r === merge.startRow && c === merge.startCol) {
                  this.tableDataWithMerged[r][c].value = startValue
                  this.tableDataWithMerged[r][c].colspan = merge.endCol - merge.startCol + 1
                } else {
                  this.tableDataWithMerged[r][c] = { value: null, merged: true, colspan: 0 } // 占位符
                }
                processedCells[r][c] = true
                if (r === merge.startRow) {
                  cellOffsets[r][c] = merge.endCol - merge.startCol
                }
              }
            }
          }
        })
      })

      // 移除已经处理过的占位符
      this.tableDataWithMerged.forEach((row, rowIndex) => {
        const newRow = []
        for (let cellIndex = 0; cellIndex < row.length; cellIndex++) {
          if (!processedCells[rowIndex][cellIndex] || (row[cellIndex].merged && row[cellIndex].value !== null)) {
            newRow.push(row[cellIndex])
          }
        }
        this.tableDataWithMerged[rowIndex] = newRow
      })

      // 处理行合并单元格的情况
      this.tableDataWithMerged.forEach((row, rowIndex) => {
        this.adjustRowForMergedCells(row, rowIndex, cellOffsets)
      })
    },

    adjustRowForMergedCells (row, rowIndex, cellOffsets) {
      const newRow = []
      let currentCol = 0

      for (let cellIndex = 0; cellIndex < row.length; cellIndex++) {
        const cell = row[cellIndex]
        if (cell.merged) {
          const colspan = this.getMergedColspan(rowIndex, cellIndex)
          const offset = cellOffsets[rowIndex][cellIndex]
          newRow.push({ ...cell, colspan, offset })
          currentCol += colspan
          const n = colspan - 1

          // 检查并调整下一个合并单元格的开始位置
          const nextMergeStarts = this.mergedCells.filter(merge => merge.startRow === rowIndex && merge.startCol > cellIndex)

          if (nextMergeStarts.length > 0 && offset > 0) {
            // console.log('查看merge', this.mergedCells)
            nextMergeStarts.forEach(nextMergeStart => {
              // console.log('offset', offset)
              // console.log('查看nextMergeStarts', nextMergeStarts)
              const nextMergeStartIndex = nextMergeStart.startCol
              const nextMergeStartCell = row[nextMergeStartIndex]
              // console.log('查看cellIndex', cellIndex)

              // 如果下一个单元格是合并单元格，调整它的开始位置
              console.log('nextMergeStartIndex', nextMergeStartIndex)
              // console.log('查看nextMergeStartCell', nextMergeStartCell)

              if (nextMergeStartCell && nextMergeStartCell.merged) {
                // console.log('cellIndex', cellIndex)
                // console.log('rowIndex', rowIndex)
                // console.log('next', row[cellIndex + 1])
                const newStartCol = cellIndex // 提前 n 列
                const newEndCol = newStartCol + (nextMergeStart.endCol - nextMergeStart.startCol)
                // console.log('newStartCol', newStartCol)
                // console.log('newEndCol', newEndCol)
                this.mergedCells = this.mergedCells.map(merge => {
                  return { ...merge, startCol: newStartCol - offset, endCol: newEndCol - offset }
                })
                row.splice(nextMergeStartIndex, 1)
                row.splice(cellIndex + 1, 0, nextMergeStartCell)
              }
            })

            // 调整当前行中所有后续合并单元格的列
            for (let i = cellIndex + 1; i < row.length; i++) {
              // console.log('查看这是第i', i)

              const nextCell = row[i]
              // console.log('nextCell.merged', nextCell.merged)
              if (nextCell.merged) {
                // console.log('查看rowIndex', rowIndex)
                // console.log('查看i', i)
                // console.log('查看mergedCells ', this.mergedCells)
                const il = i + n
                const nextMerge = this.mergedCells.find(merge => merge.startRow === rowIndex && merge.startCol === il)

                // console.log('nextMerge', nextMerge)
                // console.log('n', n)
                if (nextMerge) {
                  nextMerge.startCol -= n
                  nextMerge.endCol -= n
                  this.mergedCells = this.mergedCells.map(merge => {
                    if (merge.startRow === rowIndex && merge.startCol === i) {
                      return { ...merge, startCol: nextMerge.startCol, endCol: nextMerge.endCol }
                    }
                    return merge
                  })
                }
              }
            }
          }
        } else {
          newRow.push(cell)
          currentCol += 1
        }
      }
      this.tableDataWithMerged[rowIndex] = newRow
    },
    adjustTableColumns () {
      if (!this.tableDataWithMerged || !this.tableDataWithMerged.length) {
        return
      }

      // 计算第一行的实际列数
      const firstRowColumns = this.tableDataWithMerged[0].reduce((sum, cell, index) => sum + (cell.merged ? this.getMergedColspan(0, index) : 1), 0)
      this.tableDataWithMerged.forEach((row, rowIndex) => {
        // 计算当前行的实际列数
        const currentRowColumns = row.reduce((sum, cell, index) => sum + (cell.merged ? this.getMergedColspan(rowIndex, index) : 1), 0)
        // 如果当前行的列数与第一行不同，调整列数
        if (currentRowColumns !== firstRowColumns) {
          const newRow = []
          let currentCol = 0
          row.forEach((cell, cellIndex) => {
            if (!cell.merged || (cell.merged && cell.value !== null)) {
              newRow.push(cell)
              currentCol += cell.merged ? this.getMergedColspan(rowIndex, cellIndex) : 1
            }
          })

          // 移除多余的列
          while (newRow.length > 0 && newRow[newRow.length - 1].value === null && newRow[newRow.length - 1].merged) {
            newRow.pop()
          }

          this.tableDataWithMerged[rowIndex] = newRow
        }
      })
    },
    checkColumnsForData () {
      const columnCount = this.headers.length
      this.columnsWithData = Array(columnCount).fill(true)

      for (let colIndex = 0; colIndex < columnCount; colIndex++) {
        let hasData = false
        for (let rowIndex = 0; rowIndex < this.tableDataWithMerged.length; rowIndex++) {
          const cell = this.tableDataWithMerged[rowIndex][colIndex]
          if (cell && cell.value !== null && cell.value !== '') {
            hasData = true
            break
          }
        }
        this.columnsWithData[colIndex] = hasData
      }
    },
    getHeaderStyle (index) {
      return this.columnsWithData[index] ? { backgroundColor: '#f2f2f2' } : {}
    },
    formatCellValue (value) {
      if (typeof value === 'string' && value.includes('T')) {
        const date = new Date(value)
        if (!isNaN(date.getTime())) {
          return date.toLocaleString() // 转换为本地时间格式
        } else {
          return value // 如果转换失败，返回原始值
        }
      }
      return value
    },
    getBorderStyle (style) {
      const borderStyles = {
        thin: '1px solid #000',
        medium: '2px solid #000',
        dashed: '1px dashed #000',
        dotted: '1px dotted #000',
        double: '3px double #000',
        thick: '3px solid #000',
        hair: '1px solid #000',
        mediumDashed: '2px dashed #000',
        dashDot: '1px dashed #000',
        mediumDashDot: '2px dashed #000',
        dashDotDot: '1px dashed #000',
        mediumDashDotDot: '2px dashed #000',
        slantDashDot: '1px dashed #000'
      }

      return borderStyles[style] || '1px solid #000'
    }
  },
  mounted () {
    // 在组件挂载后，确保每行的列数与第一行的列数相同
    this.$nextTick(() => {
      this.adjustTableColumns()
      this.checkColumnsForData()
      this.forceRerender() // 强制重新渲染
    })
  }
}
</script>

<style scoped>
table {
  width: 100%;
  border-collapse: collapse;
}

th, td {
  border: 1px solid #ddd;
  padding: 8px;
  text-align: left; /* 默认左对齐，合并单元格会覆盖此样式 */
}

th {
  background-color: #f2f2f2;
  text-align: center;
}

input {
  cursor: pointer;
  margin: 25px 0;
}

.merged-cell {
  vertical-align: middle; /* 垂直居中 */
}

.text-center {
  text-align: center; /* 水平居中 */
}

.progress-bar {
  width: 100%;
  height: 20px;
  background-color: #f3f3f3;
  border-radius: 5px;
  overflow: hidden;
  margin: 10px 0;
}

.progress {
  height: 100%;
  background-color: #4caf50;
  width: 0;
  transition: width 0.3s;
}

.error-message {
  color: red;
  font-size: 14px;
  margin: 10px 0;
}

.success-message {
  color: green;
  font-size: 14px;
  margin: 10px 0;
}
</style>
