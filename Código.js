var SS = SpreadsheetApp.getActiveSpreadsheet();

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sistema de Control de Asistencia')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function buscarEmpleadoPorDni(dni) {
  var sheet = SS.getSheetByName('Empleados');
  if (!sheet) return { encontrado: false };
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(dni).trim()) {
      return {
        encontrado: true,
        dni: data[i][0],
        nombre: data[i][1],
        cargo: data[i][2],
        foto: data[i][3] || "https://i.imgur.com/6VBx3io.png"
      };
    }
  }
  return { encontrado: false };
}

function registrarMarcacion(dni, tipo) {
  var emp = buscarEmpleadoPorDni(dni);
  if (!emp.encontrado) {
    return { success: false, message: "DNI no registrado en el sistema." };
  }

  var sheetAsistencias = SS.getSheetByName('Asistencias');
  
  var now = new Date();
  var tz = Session.getScriptTimeZone();
  var fechaStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");
  var horaStr = Utilities.formatDate(now, tz, "HH:mm:ss");
  var fechaHoraStr = fechaStr + " " + horaStr;

  var estado = "SALIDA REGULAR";
  if (tipo === "Entrada") {
    var horaLimiteMinutos = (7 * 60) + 35; // 07:35 a.m.
    var partesHora = horaStr.split(":");
    var horaActualMinutos = (parseInt(partesHora[0], 10) * 60) + parseInt(partesHora[1], 10);
    estado = (horaActualMinutos > horaLimiteMinutos) ? "TARDANZA" : "PUNTUAL";
  }

  var idUnico = "REG-" + new Date().getTime();
  sheetAsistencias.appendRow([idUnico, emp.dni, emp.nombre, emp.cargo, fechaHoraStr, tipo, estado]);

  return { 
    success: true, 
    message: `¡Marcación de ${tipo} registrada correctamente!`, 
    nombre: emp.nombre,
    estado: estado 
  };
}

function verificarPasswordAdmin(passIngresada) {
  var sheetConfig = SS.getSheetByName('Config');
  var passReal = "Admin123*";
  
  if (sheetConfig) {
    var data = sheetConfig.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "PasswordAdmin") {
        passReal = String(data[i][1]).trim();
        break;
      }
    }
  }
  
  return String(passIngresada).trim() === passReal;
}

function getDashboardData() {
  var sheetEmp = SS.getSheetByName('Empleados');
  var sheetAsist = SS.getSheetByName('Asistencias');
  
  var totalEmpleados = sheetEmp ? Math.max(0, sheetEmp.getLastRow() - 1) : 0;
  
  var asistData = sheetAsist ? sheetAsist.getDataRange().getValues() : [];
  var tz = Session.getScriptTimeZone();
  var hoyStr = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");
  
  var asistenciasHoy = 0;
  var tardanzasHoy = 0;
  var registros = [];

  var fotosMap = {};
  if (sheetEmp) {
    var empData = sheetEmp.getDataRange().getValues();
    for (var k = 1; k < empData.length; k++) {
      fotosMap[String(empData[k][0]).trim()] = empData[k][3] || "https://i.imgur.com/6VBx3io.png";
    }
  }

  for (var i = 1; i < asistData.length; i++) {
    var fila = asistData[i];
    var dniEmp = String(fila[1]).trim();
    var valFecha = fila[4]; 
    
    var fechaPart = "";
    var horaPart = "";

    if (valFecha instanceof Date) {
      fechaPart = Utilities.formatDate(valFecha, tz, "dd/MM/yyyy");
      horaPart = Utilities.formatDate(valFecha, tz, "HH:mm:ss");
    } else {
      var strVal = String(valFecha).trim();
      if (strVal.indexOf(" ") !== -1) {
        var partes = strVal.split(" ");
        fechaPart = partes[0];
        horaPart = partes[1];
      } else {
        fechaPart = strVal;
        horaPart = "--:--:--";
      }
    }

    if (fechaPart === hoyStr && fila[5] === "Entrada") {
      asistenciasHoy++;
    }
    if (fechaPart === hoyStr && String(fila[6]).toUpperCase().indexOf("TARDANZA") !== -1) {
      tardanzasHoy++;
    }

    registros.push({
      id: fila[0],
      dni: dniEmp,
      nombre: fila[2],
      area: fila[3],
      fecha: fechaPart,
      hora: horaPart,
      tipo: fila[5],
      estado: fila[6],
      foto: fotosMap[dniEmp] || "https://i.imgur.com/6VBx3io.png"
    });
  }

  var empleadosLista = [];
  if (sheetEmp) {
    var empDataList = sheetEmp.getDataRange().getValues();
    for (var m = 1; m < empDataList.length; m++) {
      empleadosLista.push({
        dni: empDataList[m][0],
        nombre: empDataList[m][1],
        cargo: empDataList[m][2],
        foto: empDataList[m][3] || "https://i.imgur.com/6VBx3io.png"
      });
    }
  }

  return {
    totalEmpleados: totalEmpleados,
    asistenciasHoy: asistenciasHoy,
    tardanzasHoy: tardanzasHoy,
    registros: registros.reverse(),
    empleados: empleadosLista
  };
}

function guardarEmpleado(dni, nombre, cargo, foto) {
  var sheet = SS.getSheetByName('Empleados');
  if (!sheet) return { success: false, message: "No se encontró la hoja Empleados" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(dni).trim()) {
      sheet.getRange(i + 1, 2).setValue(nombre);
      sheet.getRange(i + 1, 3).setValue(cargo);
      if(foto) {
        sheet.getRange(i + 1, 4).setValue(foto);
      }
      return { success: true, message: "Empleado actualizado correctamente." };
    }
  }
  sheet.appendRow([dni, nombre, cargo, foto || "https://i.imgur.com/6VBx3io.png"]);
  return { success: true, message: "Empleado registrado con éxito." };
}

function eliminarEmpleado(dni) {
  var sheet = SS.getSheetByName('Empleados');
  if (!sheet) return { success: false, message: "No se encontró la hoja Empleados" };
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(dni).trim()) {
      sheet.deleteRow(i + 1);
      return { success: true, message: "Empleado eliminado correctamente." };
    }
  }
  return { success: false, message: "No se encontró el DNI especificado." };
}