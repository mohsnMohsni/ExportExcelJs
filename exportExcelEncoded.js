function utf8_to_b64( str ) {
  return window.btoa(unescape(encodeURIComponent( str )));
}

function b64_to_utf8( str ) {
  return decodeURIComponent(escape(window.atob( str )));
}

let uri = "data:application/vnd.ms-excel;base64,",
  template = '<html><head><meta charset="UTF-8"></head><body><table>{table}</table></body></html>',
  base64 = function (s) {
    return window.btoa(unescape(encodeURIComponent(s)));
  },
  format = function (s, c) {
    return s.replace(/{(\w+)}/g, function (m, p) {
      return c[p];
    });
  };

table = document.getElementById("excelData");
let ctx = { worksheet: null || "Worksheet", table: table.innerHTML };

let a = document.createElement("a");
a.href = uri + window.btoa(unescape(encodeURIComponent("\uFEFF" + format(template, ctx))));
a.download = `discountcode ${new Date().toLocaleDateString("fa-IR")}.xls`;
a.click();
