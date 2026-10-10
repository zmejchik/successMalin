<template>
  <div class="card">
    <h1 class="card-title">Створення динамічної таблиці даних</h1>
    <p class="card-subtitle">Завантаж і не парся щас все буде, а може і ні</p>

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
          {{ file ? file.name : 'Выберите файл или перетащите его сюда' }}
        </span>
        <span class="drop-zone__hint">Поддерживаются форматы .xlsx, .xls, .xlsm, .csv</span>
      </label>

      <div class="actions">
        <button class="btn btn--primary" @click="processFile" :disabled="!file || isProcessing">
          {{ isProcessing ? 'Обробка…' : 'Розрахувати' }}
        </button>

        <button v-if="file && !isProcessing" class="btn btn--text" @click="reset" type="button">
          Сбросить
        </button>
      </div>
    </div>

    <div v-if="message" class="snackbar" :class="{ 'snackbar--error': isError }" role="status">
      {{ message }}
    </div>
  </div>
</template>

<script>
import * as XLSX from 'xlsx'
import { saveAs } from 'file-saver'

// ---------------------------------------------------------------------------
// Разбирает значение из столбца I в { base, num }.
//   "Q"    → { base: 'Q', num: '' }
//   "Q1"   → { base: 'Q', num: '1' }
//   "Q 1"  → { base: 'Q', num: '1' }
//   "Q-2"  → { base: 'Q', num: '2' }
//   "1Q"   → { base: 'Q', num: '1' }
//   "Щось" → { base: 'Щ', num: '' }
// ---------------------------------------------------------------------------
function parseType(raw) {
  if (raw === '' || raw === null || raw === undefined) return null
  const s = String(raw).trim().toUpperCase()
  if (!s) return null

  // буквы [разделитель] цифры → Q1, Q 1, Q-1, Q.1, Q_1
  let m = s.match(/^([A-ZА-ЯЇІЄҐ]+)\s*[-_.]?\s*(\d+)/)
  if (m) return { base: m[1], num: m[2] }

  // цифры [разделитель] буквы → 1Q, 1 Q, 1-Q
  m = s.match(/^(\d+)\s*[-_.]?\s*([A-ZА-ЯЇІЄҐ]+)/)
  if (m) return { base: m[2], num: m[1] }

  // только буквы → Q, S, W, M, P, G
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

        this.message = 'Таблица успешно создана. Файл скачан.'
        this.isError = false
      } catch (err) {
        console.error(err)
        this.message = 'Ошибка: ' + err.message
        this.isError = true
      } finally {
        this.isProcessing = false
      }
    },

    createDynamicAverageTable(workbook, sheetName, data) {
      const sheet = workbook.Sheets[sheetName]

      // --- 1. lastRow по столбцу F (idx 5) ---
      let lastRow = 0
      for (let r = data.length - 1; r >= 0; r--) {
        const v = data[r]?.[5]
        if (v !== '' && v !== null && v !== undefined) {
          lastRow = r
          break
        }
      }

      if (lastRow < 1) {
        alert('No data')
        return
      }

      // --- 2. Студенты + множество базовых букв и полных ключей ---
      const students = []
      const studentSet = new Set()
      const baseSet = new Set() // только буквы: Q, W, G, S…
      const fullSet = new Set() // буква+цифра: Q1, Q2, W1…

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

      // --- 3. Формируем список колонок ---
      // Сначала ВСЕ общие (по букве), потом ВСЕ детальные (буква+номер).
      //
      // Каждая колонка — это объект:
      //   { base, num, matchNum, label }
      // где:
      //   matchNum === null → учитываем любые строки с этой base-буквой
      //   matchNum === '1'  → учитываем только строки с base + 1
      const columns = []

      // 3a. Общие колонки: только буква, без номера
      Array.from(baseSet)
        .sort()
        .forEach((base) => {
          columns.push({
            base,
            num: '',
            matchNum: null,
            label: baseLabelOf(base),
          })
        })

      // 3b. Детальные колонки: буква + номер, сортируем по букве, потом по номеру
      Array.from(fullSet)
        .sort((a, b) => {
          const ma = a.match(/^([A-ZА-ЯЇІЄҐ]+)(\d+)$/)
          const mb = b.match(/^([A-ZА-ЯЇІЄҐ]+)(\d+)$/)
          if (!ma || !mb) return a.localeCompare(b)
          if (ma[1] !== mb[1]) return ma[1].localeCompare(mb[1])
          return parseInt(ma[2], 10) - parseInt(mb[2], 10)
        })
        .forEach((key) => {
          const m = key.match(/^([A-ZА-ЯЇІЄҐ]+)(\d+)$/)
          const base = m[1]
          const num = m[2]
          columns.push({
            base,
            num,
            matchNum: num,
            label: baseLabelOf(base) + num,
          })
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

      // --- 5. Запись ячейки ---
      const setCell = (r, c, value) => {
        const addr = XLSX.utils.encode_cell({ r, c })
        if (value === '' || value === null || value === undefined) {
          delete sheet[addr]
          return
        }
        const t = typeof value === 'number' ? 'n' : 's'
        sheet[addr] = { t, v: value }
      }

      // --- 6. Заголовки ---
      setCell(0, K_COL, 'Student')
      columns.forEach((col, idx) => {
        setCell(0, K_COL + 1 + idx, col.label)
      })

      // --- 7. Заполняем таблицу ---
      const vbaRound = (x) => {
        const floor = Math.floor(x)
        const diff = x - floor
        if (diff > 0.5) return floor + 1
        if (diff < 0.5) return floor
        return floor % 2 === 0 ? floor : floor + 1
      }

      students.forEach((student, sIdx) => {
        const rowOut = sIdx + 1
        setCell(rowOut, K_COL, student)

        columns.forEach((col, cIdx) => {
          let sumVal = 0
          let countVal = 0

          for (let i = 1; i <= lastRow; i++) {
            const nameI = data[i]?.[5]
            const typeI = data[i]?.[8]
            const valueH = data[i]?.[7]

            if (String(nameI) !== student) continue

            const parsed = parseType(typeI)
            if (!parsed) continue
            if (parsed.base !== col.base) continue

            // Для общих колонок (matchNum === null) — учитываем любые номера.
            // Для детальных — только точное совпадение base+num.
            if (col.matchNum !== null && parsed.num !== col.matchNum) continue

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

          if (countVal === 0) {
            setCell(rowOut, K_COL + 1 + cIdx, '')
          } else {
            setCell(rowOut, K_COL + 1 + cIdx, vbaRound(sumVal / countVal))
          }
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
}

.btn--primary {
  background-color: #f66f4d;
  color: #ffffff;
  box-shadow: 0px 4px 12px rgba(246, 111, 77, 0.25);
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
}

.snackbar--error {
  background-color: #fce8e6;
  color: #c5221f;
  border-color: #f6aea9;
}

/* ---------- Адаптив ---------- */
@media (max-width: 640px) {
  .card {
    padding: 24px 20px;
    border-radius: 8px;
  }

  .card-title {
    font-size: 24px;
    line-height: 32px;
  }

  .actions {
    flex-direction: column;
    align-items: stretch;
  }

  .btn--text {
    height: 40px;
  }
}
</style>
