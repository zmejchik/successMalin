<template>
  <div class="wrapper">
    <header class="navbar">
      <div class="navbar-logo">
        <img src="/logo.svg" alt="Logo" class="navbar-logo-img" />
        <span class="navbar-logo-text">Success school!</span>
      </div>
    </header>

    <main class="page">
      <ExcelAverageProcessor />
    </main>
  </div>
</template>

<script setup>
import ExcelAverageProcessor from './components/ExcelAverageProcessor.vue'
</script>

<style scoped>
.wrapper {
  max-width: 1156px;
  margin: 24px auto 60px;
  padding: 0 20px;
}

/* ---------- Навбар ---------- */
.navbar {
  height: 54px;
  display: flex;
  flex-direction: row;
  align-items: center;
  margin-bottom: 60px;
}

.navbar-logo {
  display: inline-flex;
  align-items: center;
  gap: 10px;
  min-width: 0;
}

.navbar-logo-img {
  width: 40px;
  height: 40px;
  object-fit: contain;
  flex-shrink: 0;
}

.navbar-logo-text {
  font-family: 'Sen', sans-serif;
  font-weight: 700;
  font-size: 20px;
  color: #2d3134;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.menu {
  margin-left: 70px;
  display: flex;
  gap: 36px;
}

.linkmenu {
  text-decoration: none;
  color: #2e476b;
  font-family: 'Poppins', sans-serif;
  font-size: 16px;
  transition: color 0.2s ease;
}

.linkmenu:hover {
  color: #f66f4d;
}

/* ---------- Страница ---------- */
.page {
  display: flex;
  justify-content: center;
}

/* ---------- Планшет (до 900px) ---------- */
@media (max-width: 900px) {
  .wrapper {
    padding: 0 16px;
    margin: 16px auto 40px;
  }

  .navbar {
    margin-bottom: 40px;
  }
}

/* ---------- Мобильные (до 640px) ---------- */
@media (max-width: 640px) {
  .wrapper {
    padding: 0 12px;
    margin: 12px auto 32px;
  }

  .navbar {
    height: auto;
    flex-direction: row;
    align-items: center;
    justify-content: center;
    margin-bottom: 24px;
    padding: 12px 0;
    border-bottom: 1px solid #f0ebe2;
  }

  .navbar-logo {
    gap: 8px;
    width: 100%;
    justify-content: center;
  }

  .navbar-logo-img {
    width: 32px;
    height: 32px;
  }

  .navbar-logo-text {
    font-size: 16px;
  }

  .menu {
    margin-left: 0;
    gap: 20px;
  }
}

/* ---------- Очень узкие экраны (до 380px) ---------- */
@media (max-width: 380px) {
  .navbar-logo-text {
    font-size: 14px;
    white-space: normal;
    text-align: center;
    line-height: 1.2;
  }
}
</style>
