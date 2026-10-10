<template>
  <div class="wrapper">
    <header class="navbar">
      <div class="navbar-logo">
        <img src="/logo.svg" alt="Logo" class="navbar-logo-img" />
        <span class="navbar-logo-text">Malinovska Nataliia forever</span>
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

@media (max-width: 768px) {
  .navbar {
    flex-direction: column;
    height: auto;
    align-items: flex-start;
    gap: 16px;
    margin-bottom: 32px;
  }

  .menu {
    margin-left: 0;
    gap: 20px;
  }
}
</style>
