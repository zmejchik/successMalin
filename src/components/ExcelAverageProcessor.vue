<template>
  <div class="card">
    <h1 class="card-title">Tworzenie dynamicznej tabeli danych</h1>
    <p class="card-subtitle">Załaduj plik i nie przejmuj się — wszystko będzie, a może i nie</p>

    <div class="uploader">
      <label class="drop-zone" :class="{ 'has-file': file, disabled: isProcessing }">
        <input
          type="file"
          accept=".xlsx,.xls,.xlsm,.csv"
          @change="onFileSelected"
          :disabled="isProcessing"
        />
        <svg class="drop-zone__icon" viewBox="0 0 24 24" width="44" height="44" aria-hidden="true">
          <path
            fill="#F66F4D"
            d="M19 15v4H5v-4H3v4a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-4h-2ZM11 5.83 8.41 8.41 7 7l5-5 5 5-1.41 1.41L13 5.83V16h-2V5.83Z"
          />
        </svg>
        <span class="drop-zone__title">
          {{ file ? file.name : 'Wybierz plik lub przeciągnij go tutaj' }}
        </span>
        <span class="drop-zone__hint">Obsługiwane formaty: .xlsx, .xls, .xlsm, .csv</span>
      </label>

      <div class="actions">
        <button class="btn btn--primary" @click="processFile" :disabled="!file || isProcessing">
          {{ isProcessing ? 'Przetwarzanie…' : 'Oblicz' }}
        </button>

        <button v-if="file && !isProcessing" class="btn btn--text" @click="reset" type="button">
          Wyczyść
        </button>
      </div>
    </div>

    <div v-if="message" class="snackbar" :class="{ 'snackbar--error': isError }" role="status">
      {{ message }}
    </div>
  </div>
</template>

<script>
import * as XLSX from 'xlsx-js-style'
import { saveAs } from 'file-saver'

// ---------------------------------------------------------------------------
// Парсинг типа из столбца I
// ---------------------------------------------------------------------------
function parseType(raw) {
  if (raw === '' || raw === null || raw === undefined) return null
  const s = String(raw).trim().toUpperCase()
  if (!s) return null

  let m = s.match(/^([A-ZА-ЯЇІЄҐ]+)\s*[-_.]?\s*(\d+)/)
  if (m) return { base: m[1], num: m[2] }

  m = s.match(/^(\d+)\s*[-_.]?\s*([A-ZА-ЯЇІЄҐ]+)/)
  if (m) return { base: m[2], num: m[1] }

  m = s.match(/^([A-ZА-ЯЇІЄҐ]+)/)
  if (m) return { base: m[1], num: '' }

  return null
}

const BASE_LABELS = {
  S: 'Mówienie S',
  W: 'Pisanie W',
  M: 'M',
  P: 'P',
  Q: 'Słownictwo Q',
  G: 'Gramatyka G',
}

function baseLabelOf(base) {
  return BASE_LABELS[base] ?? base
}

// ---------- Стили ----------
const YELLOW = { patternType: 'solid', fgColor: { rgb: 'FFF9C4' } }
const GREEN = { patternType: 'solid', fgColor: { rgb: 'C8E6C9' } }

const STYLE_HEADER_DEFAULT = {
  font: { bold: true, sz: 12 },
  alignment: { textRotation: 90, vertical: 'bottom', horizontal: 'center' },
}

const STYLE_HEADER_YELLOW = {
  font: { bold: true, sz: 12 },
  fill: YELLOW,
  alignment: { textRotation: 90, vertical: 'bottom', horizontal: 'center' },
}

const STYLE_HEADER_GREEN = {
  font: { bold: true, sz: 12 },
  fill: GREEN,
  alignment: { textRotation: 90, vertical: 'bottom', horizontal: 'center' },
}

const STYLE_BODY_DEFAULT = {}

const STYLE_BODY_YELLOW = {
  fill: YELLOW,
}

const STYLE_BODY_GREEN = {
  fill: GREEN,
}

export default {
  name: 'ExcelAverageProcessor',
  data() {
    return {
      file: null,
      isProcessing: false,
      message: '',
      isError: false,
    }
  },
  methods: {
    onFileSelected(event) {
      this.file = event.target.files[0] || null
      this.message = ''
    },

    reset() {
      this.file = null
      this.message = ''
      this.isError = false
    },

    async processFile() {
      if (!this.file) return
      this.isProcessing = true
      this.message = ''

      try {
        const arrayBuffer = await this.file.arrayBuffer()
        const workbook = XLSX.read(arrayBuffer, { type: 'array' })

        const sheetName = workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]

        const data = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
          raw: true,
        })

        this.createDynamicAverageTable(workbook, sheetName, data)

        const out = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' })
        const blob = new Blob([out], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        })
        saveAs(blob, `average_${this.file.name.replace(/\.[^.]+$/, '')}.xlsx`)

        this.message = 'Tabela została pomyślnie utworzona. Plik został pobrany.'
        this.isError = false
      } catch (err) {
        console.error(err)
        this.message = 'Błąd: ' + err.message
        this.isError = true
      } finally {
        this.isProcessing = false
      }
    },

    createDynamicAverageTable(workbook, sheetName, data) {
      const sheet = workbook.Sheets[sheetName]

      // --- 1. lastRow по столбцу F ---
      let lastRow = 0
      for (let r = data.length - 1; r >= 0; r--) {
        const v = data[r]?.[5]
        if (v !== '' && v !== null && v !== undefined) {
          lastRow = r
          break
        }
      }

      if (lastRow < 1) {
        alert('Brak danych')
        return
      }

      // --- 2. Студенты, базовые буквы, полные ключи ---
      const students = []
      const studentSet = new Set()
      const baseSet = new Set()
      const fullSet = new Set()

      for (let i = 1; i <= lastRow; i++) {
        const name = data[i]?.[5] ?? ''
        const t = data[i]?.[8] ?? ''

        if (name !== '' && name !== null && name !== undefined) {
          const key = String(name)
          if (!studentSet.has(key)) {
            studentSet.add(key)
            students.push(key)
          }
        }

        const parsed = parseType(t)
        if (parsed) {
          baseSet.add(parsed.base)
          if (parsed.num) fullSet.add(parsed.base + parsed.num)
        }
      }

      // --- 3. Формируем колонки ---
      const columns = []

      // 3a. Общие
      Array.from(baseSet)
        .sort()
        .forEach((base) => {
          columns.push({
            type: 'base',
            base,
            matchNum: null,
            label: baseLabelOf(base),
          })
        })

      // 3b. Группы по номерам
      const numSet = new Set()
      Array.from(fullSet).forEach((key) => {
        const m = key.match(/^([A-ZА-ЯЇІЄҐ]+)(\d+)$/)
        if (m) numSet.add(m[2])
      })

      Array.from(numSet)
        .sort((a, b) => parseInt(a, 10) - parseInt(b, 10))
        .forEach((num) => {
          const basesWithNum = Array.from(baseSet)
            .filter((base) => fullSet.has(base + num))
            .sort()

          // Детальные колонки
          basesWithNum.forEach((base) => {
            columns.push({
              type: 'detail',
              base,
              matchNum: num,
              label: baseLabelOf(base) + num,
            })
          })

          // Колонка среднего для группы
          if (basesWithNum.length > 0) {
            columns.push({
              type: 'average',
              num,
              bases: basesWithNum,
              label: `Średnia ${num}`,
            })
          }
        })

      // --- 4. Очистка K1:ZZ10000 ---
      const K_COL = 10
      const ZZ_COL = 701
      const CLEAR_MAX_ROW = 10000

      for (let r = 0; r < CLEAR_MAX_ROW; r++) {
        for (let c = K_COL; c <= ZZ_COL; c++) {
          const addr = XLSX.utils.encode_cell({ r, c })
          delete sheet[addr]
        }
      }

      // --- 5. Запись ячейки со стилем ---
      const setCell = (r, c, value, style) => {
        const addr = XLSX.utils.encode_cell({ r, c })
        if (value === '' || value === null || value === undefined) {
          if (style && style.fill) {
            sheet[addr] = { t: 's', v: '', s: style }
          } else {
            delete sheet[addr]
          }
          return
        }
        const t = typeof value === 'number' ? 'n' : 's'
        sheet[addr] = { t, v: value, s: style || STYLE_BODY_DEFAULT }
      }

      // --- 6. Заголовки ---
      const headerStyleFor = (col) => {
        if (col.type === 'average') return STYLE_HEADER_GREEN
        if (col.type === 'detail') return STYLE_HEADER_YELLOW
        return STYLE_HEADER_DEFAULT
      }

      setCell(0, K_COL, 'Student', STYLE_HEADER_DEFAULT)
      columns.forEach((col, idx) => {
        setCell(0, K_COL + 1 + idx, col.label, headerStyleFor(col))
      })

      // --- 7. Заполняем таблицу ---
      const vbaRound = (x) => {
        const floor = Math.floor(x)
        const diff = x - floor
        if (diff > 0.5) return floor + 1
        if (diff < 0.5) return floor
        return floor % 2 === 0 ? floor : floor + 1
      }

      const rowMatchesColumn = (col, parsed) => {
        if (!parsed) return false
        if (col.type === 'average') {
          return parsed.num === col.num && col.bases.includes(parsed.base)
        }
        if (col.type === 'detail') {
          return parsed.base === col.base && parsed.num === col.matchNum
        }
        return parsed.base === col.base
      }

      const bodyStyleFor = (col) => {
        if (col.type === 'average') return STYLE_BODY_GREEN
        if (col.type === 'detail') return STYLE_BODY_YELLOW
        return STYLE_BODY_DEFAULT
      }

      students.forEach((student, sIdx) => {
        const rowOut = sIdx + 1
        setCell(rowOut, K_COL, student, STYLE_BODY_DEFAULT)

        // ★★ ПЕРВЫЙ ПРОХОД: общие и детальные колонки.
        // Запоминаем посчитанные детальные значения в detailValues
        // вида { 'G1': 4, 'Q1': 5, 'W1': '', ... }
        const detailValues = {}

        columns.forEach((col, cIdx) => {
          if (col.type === 'average') return // средние — во втором проходе

          let sumVal = 0
          let countVal = 0

          for (let i = 1; i <= lastRow; i++) {
            const nameI = data[i]?.[5]
            if (String(nameI) !== student) continue

            const parsed = parseType(data[i]?.[8])
            if (!rowMatchesColumn(col, parsed)) continue

            const valueH = data[i]?.[7]
            if (
              valueH !== '' &&
              valueH !== null &&
              valueH !== undefined &&
              valueH !== ' ' &&
              !isNaN(Number(valueH))
            ) {
              sumVal += Number(valueH)
              countVal += 1
            }
          }

          const style = bodyStyleFor(col)
          const value = countVal === 0 ? '' : vbaRound(sumVal / countVal)
          setCell(rowOut, K_COL + 1 + cIdx, value, style)

          // Для детальных колонок — запоминаем значение под ключом "G1", "Q1" и т.д.
          if (col.type === 'detail') {
            detailValues[col.base + col.matchNum] = value
          }
        })

        // ★★ ВТОРОЙ ПРОХОД: колонки Średnia N.
        // Берём уже посчитанные средние по каждой букве группы и усредняем их.
        columns.forEach((col, cIdx) => {
          if (col.type !== 'average') return

          let sumAvg = 0
          let countAvg = 0

          col.bases.forEach((base) => {
            const v = detailValues[base + col.num]
            // Пустые (нет данных по этой букве) — НЕ учитываем
            if (v !== '' && v !== null && v !== undefined && !isNaN(Number(v))) {
              sumAvg += Number(v)
              countAvg += 1
            }
          })

          const value = countAvg === 0 ? '' : vbaRound(sumAvg / countAvg)
          setCell(rowOut, K_COL + 1 + cIdx, value, STYLE_BODY_GREEN)
        })
      })

      // --- 8. Обновляем !ref ---
      const range = XLSX.utils.decode_range(sheet['!ref'] || 'A1')
      range.e.r = Math.max(range.e.r, CLEAR_MAX_ROW - 1)
      range.e.c = Math.max(range.e.c, ZZ_COL)
      sheet['!ref'] = XLSX.utils.encode_range(range)

      workbook.Sheets[sheetName] = sheet
    },
  },
}
</script>

<style scoped>
/* ---------- Карточка ---------- */
.card {
  width: 100%;
  max-width: 720px;
  background-color: #ffffff;
  border-radius: 10px;
  padding: 40px 40px 32px;
  box-shadow: 0px 2px 6px rgba(229, 137, 24, 0.08);
}

.card-title {
  font-family: 'Sen', sans-serif;
  font-weight: 700;
  font-size: 32px;
  line-height: 40px;
  color: #2d3134;
  margin: 0 0 12px;
  word-wrap: break-word;
}

.card-subtitle {
  font-family: 'Inter', sans-serif;
  font-weight: 400;
  font-size: 16px;
  line-height: 26px;
  color: #5b5f62;
  margin: 0 0 32px;
}

/* ---------- Зона загрузки ---------- */
.uploader {
  display: flex;
  flex-direction: column;
  gap: 20px;
}

.drop-zone {
  position: relative;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 8px;
  padding: 32px 24px;
  border: 1.5px dashed #e0dad0;
  border-radius: 10px;
  background-color: #faf8ed;
  cursor: pointer;
  text-align: center;
  transition: background 0.2s ease, border-color 0.2s ease;
  min-height: 160px;
}

.drop-zone:hover {
  background-color: #f5f0e3;
  border-color: #f66f4d;
}

.drop-zone.has-file {
  border-style: solid;
  border-color: #f66f4d;
  background-color: #fff4f0;
}

.drop-zone.disabled {
  opacity: 0.6;
  cursor: not-allowed;
}

.drop-zone input[type='file'] {
  position: absolute;
  inset: 0;
  opacity: 0;
  cursor: pointer;
}

.drop-zone__icon {
  opacity: 0.9;
  width: 44px;
  height: 44px;
  flex-shrink: 0;
}

.drop-zone__title {
  font-family: 'Inter', sans-serif;
  font-weight: 500;
  font-size: 15px;
  color: #2d3134;
  max-width: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.drop-zone__hint {
  font-family: 'Inter', sans-serif;
  font-size: 13px;
  color: #5b5f62;
}

/* ---------- Кнопки ---------- */
.actions {
  display: flex;
  gap: 12px;
  align-items: center;
  flex-wrap: wrap;
}

.btn {
  border: none;
  border-radius: 10px;
  padding: 0 28px;
  height: 48px;
  font-family: 'Inter', sans-serif;
  font-size: 15px;
  font-weight: 600;
  letter-spacing: 0.2px;
  cursor: pointer;
  transition: background 0.2s ease, box-shadow 0.2s ease, color 0.2s ease;
  -webkit-tap-highlight-color: transparent;
}

.btn--primary {
  background-color: #f66f4d;
  color: #ffffff;
  box-shadow: 0px 4px 12px rgba(246, 111, 77, 0.25);
  flex: 1 1 auto;
  min-width: 180px;
}

.btn--primary:hover:not(:disabled) {
  background-color: #e85e3b;
  box-shadow: 0px 6px 16px rgba(246, 111, 77, 0.35);
}

.btn--primary:disabled {
  background-color: #e0dad0;
  color: #9aa0a6;
  box-shadow: none;
  cursor: not-allowed;
}

.btn--text {
  background: transparent;
  color: #2e476b;
  padding: 0 16px;
  flex: 0 0 auto;
}

.btn--text:hover {
  background-color: rgba(46, 71, 107, 0.06);
  color: #f66f4d;
}

/* ---------- Уведомление ---------- */
.snackbar {
  margin-top: 24px;
  padding: 14px 18px;
  border-radius: 10px;
  font-family: 'Inter', sans-serif;
  font-size: 14px;
  background-color: #e6f4ea;
  color: #137333;
  border: 1px solid #ceead6;
  word-wrap: break-word;
}

.snackbar--error {
  background-color: #fce8e6;
  color: #c5221f;
  border-color: #f6aea9;
}

/* ---------- Планшет (до 900px) ---------- */
@media (max-width: 900px) {
  .card {
    padding: 32px 28px 26px;
  }

  .card-title {
    font-size: 26px;
    line-height: 34px;
  }
}

/* ---------- Мобильные (до 640px) ---------- */
@media (max-width: 640px) {
  .card {
    padding: 24px 20px 20px;
    border-radius: 8px;
  }

  .card-title {
    font-size: 22px;
    line-height: 28px;
    margin-bottom: 8px;
  }

  .card-subtitle {
    font-size: 14px;
    line-height: 22px;
    margin-bottom: 24px;
  }

  .uploader {
    gap: 16px;
  }

  .drop-zone {
    padding: 24px 16px;
    min-height: 140px;
  }

  .drop-zone__icon {
    width: 36px;
    height: 36px;
  }

  .drop-zone__title {
    font-size: 14px;
    white-space: normal;
    line-height: 1.35;
  }

  .drop-zone__hint {
    font-size: 12px;
  }

  .actions {
    flex-direction: column;
    align-items: stretch;
    gap: 10px;
  }

  .btn {
    width: 100%;
    height: 46px;
    padding: 0 16px;
    min-width: 0;
  }

  .btn--text {
    height: 40px;
  }

  .snackbar {
    font-size: 13px;
    padding: 12px 14px;
  }
}

/* ---------- Очень узкие экраны (до 380px) ---------- */
@media (max-width: 380px) {
  .card {
    padding: 20px 16px 16px;
  }

  .card-title {
    font-size: 19px;
    line-height: 25px;
  }

  .drop-zone {
    padding: 20px 12px;
  }
}
</style>
