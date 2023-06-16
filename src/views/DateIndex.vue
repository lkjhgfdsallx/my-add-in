<script setup>
async function onSetColor() {
  window.Office.context.ui.messageParent(
    JSON.stringify({
      type: 'setRangeColor',
      color: 'red',
    }),
  )
}
</script>

<template>
  <div>
    <div class="content">
      <div class="content-header">
        <div class="padding">
          <h1>你好世界</h1>
        </div>
      </div>
      <div class="content-main">
        <div class="padding">
          <p>选择下面的按钮，将所选范围的颜色设置为 红色。</p>
          <br>
          <h3>试试看</h3>
          <button @click="onSetColor">
            设置颜色
          </button>
        </div>
      </div>
    </div>
  </div>
</template>

<style>
.content-header {
  background: #2a8dd4;
  color: #fff;
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 80px;
  overflow: hidden;
}

.content-main {
  background: #fff;
  position: fixed;
  top: 80px;
  left: 0;
  right: 0;
  bottom: 0;
  overflow: auto;
}

.padding {
  padding: 15px;
}
</style>
