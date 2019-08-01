<template>
  <el-container>
    <el-button @click="handleClick">{{test}}</el-button>
  </el-container>
</template>

<script>

export default {
  name:'Home',
  data () {
    return {
      test:'test'
    }
  },
  methods:{
    handleClick () {
      Excel.run(async (context) => {
        let sheetName = 'Sheet1'
        let worksheet = context.workbook.worksheets.getItem(sheetName);
        let sourceRange = worksheet.getRange("B2:E2")
        sourceRange.load("format/fill/color, format/font/name, format/font/color")

        

        await context.sync()
        let targetRange = worksheet.getRange("B9:E9")
        targetRange.set(sourceRange)
        targetRange.format.autofitColumns()
        let targetRange2 = worksheet.getRange("B7:E7")
        targetRange2.set(sourceRange);
        targetRange2.format.autofitColumns()

        await context.sync()

      })
      .catch(function (error) {
          console.log("Error: " + error);
          if (error instanceof OfficeExtension.Error) {
              console.log("Debug info: " + JSON.stringify(error.debugInfo));
          }
      })
    }
  }
}
</script>

<style lang="stylus" scoped>

</style>



