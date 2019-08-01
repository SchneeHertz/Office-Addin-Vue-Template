/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as Vue from "vue"
import router from './router'
import ElementUI from 'element-ui'
import 'element-ui/lib/theme-chalk/index.css'
import _ from 'lodash'

Vue.use(ElementUI)
window._ = _

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    new Vue({
      el: "#app",
      router
    })
  }
})

// export async function run() {
//   try {
//     await Excel.run(async context => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
