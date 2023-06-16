// 参考：https://learn.microsoft.com/en-us/office/dev/add-ins/develop/dialog-api-in-office-add-ins
async function displayDialogAsync(link, width, height, title) {
  return new Promise((resolve, reject) => {
    window.Office.context.ui.displayDialogAsync(
      link,
      { width, height, title, displayInIframe: true },
      (result) => {
        if (result.status === window.Office.AsyncResultStatus.Failed) {
          reject(result.error)
        }
        else {
          const dialog = result.value
          dialog.addEventHandler(
            window.Office.EventType.DialogMessageReceived,
            (arg) => {
              const message = JSON.parse(arg.message)
              if (message.type === 'setRangeColor')
                setRangeColor(message.color)
            },
          )
          dialog.addEventHandler(window.Office.EventType.DialogEventReceived, () => {
            // 处理弹窗的事件
          })
          resolve(dialog)
        }
      },
    )
  })
}

async function insertFunction(event) {
  try {
    await displayDialogAsync('https://localhost:3000/?1ink=Insertion', 20, 40, '插入函数')
    event.completed()
  }
  catch (error) {
    console.error(error)
  }
}

async function productsFunction(event) {
  try {
    await displayDialogAsync('https://localhost:3000/?1ink=products', 20, 40)
    event.completed()
  }
  catch (error) {
    console.error(error)
  }
}

async function dateFunction(event) {
  try {
    await displayDialogAsync('https://localhost:3000/?1ink=date', 20, 40)
    event.completed()
  }
  catch (error) {
    console.error(error)
  }
}

async function flushedFunction(event) {
  try {
    await displayDialogAsync('https://localhost:3000/?1ink=flushed', 20, 40)
    event.completed()
  }
  catch (error) {
    console.error(error)
  }
}

async function helpFunction(event) {
  try {
    await displayDialogAsync('https://localhost:3000/?1ink=help', 20, 40)
    event.completed()
  }
  catch (error) {
    console.error(error)
  }
}

async function loginFunction(event) {
  try {
    await displayDialogAsync('https://localhost:3000/?1ink=login', 20, 40)
    event.completed()
  }
  catch (error) {
    console.error(error)
  }
}

async function setRangeColor(color) {
  try {
    await window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange()
      range.format.fill.color = color
      await context.sync()
    })
  }
  catch (error) {
    console.error(error)
  }
}

export { insertFunction, productsFunction, dateFunction, helpFunction, loginFunction, flushedFunction }
