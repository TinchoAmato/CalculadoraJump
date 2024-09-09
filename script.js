function obtenerDatosFormulario() {
    const nombre = document.getElementById('nombre').value.toUpperCase();
    const cumpleanero = document.getElementById('cumpleanero').value;
    const edad = parseInt(document.getElementById('edad').value);
    const telefono = document.getElementById('telefono').value;
    const saltadores = parseInt(document.getElementById('saltadores').value);
    const kids = parseInt(document.getElementById('kids').value) || 0;
    const noSaltadores = parseInt(document.getElementById('noSaltadores').value);
    const fecha = document.getElementById('fecha').value;
    const hora = document.getElementById('hora').value;
  
    if (saltadores === 0 && kids === 0) {
      alert('Debe haber al menos saltadores o kids.');
      return null;
    }
  
    return { nombre, cumpleanero, edad, telefono, saltadores, kids, noSaltadores, fecha, hora };
  }
  
  function calcularPresupuesto({ nombre, cumpleanero, edad, telefono, saltadores, kids, noSaltadores, fecha, hora }) {
    const diasSemana = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];
    const diaSemana = diasSemana[new Date(fecha).getDay() + 1]; // Obtener el día siguiente
    const [hours, minutes] = hora.split(':');
    const hour = parseInt(hours);
  
    const zona = edad < 6 ? 'Zona kids' : 'Camas Elásticas';
  
    let costoSaltadores = 0;
    let costoKids = 0;
    const costoAdultos = 12500;
  
    if (hour >= 13 && hour < 14) {
      costoSaltadores = 20500;
    } else {
      costoSaltadores = 23000;
    }
  
    if (diaSemana === 'Sábado' || diaSemana === 'Domingo') {
      costoSaltadores = 27000;
    } else if (diaSemana === 'Viernes' && hour >= 14) {
      costoSaltadores = 23000;
    }
  
    if (diaSemana === 'Viernes' || diaSemana === 'Sábado' || diaSemana === 'Domingo') {
      costoKids = 16000;
    } else {
      costoKids = 13000;
    }
  
    let adultosExcluidos = 0;
    const totalSaltadoresKids = saltadores + kids;
  
    if (totalSaltadoresKids >= 15) {
      adultosExcluidos = 4;
    } else if (totalSaltadoresKids >= 10) {
      adultosExcluidos = 2;
    }
  
    const totalSaltadores = saltadores * costoSaltadores;
    const totalKids = kids * costoKids;
    const totalAdultos = Math.max(0, noSaltadores - adultosExcluidos) * costoAdultos;
    const presupuestoTotal = totalSaltadores + totalKids + totalAdultos;
  
    const finalDate = new Date(0, 0, 0, hours, minutes);
    finalDate.setHours(finalDate.getHours() + 2, finalDate.getMinutes() + 30);
  
    const formatHour = (hours, minutes) => minutes === '00' ? hours : `${hours}:${minutes}`;
    const initialTime = formatHour(hours, minutes);
    const finalTime = formatHour(finalDate.getHours().toString().padStart(2, '0'), finalDate.getMinutes().toString().padStart(2, '0'));
    const formattedDate = new Date(new Date(fecha).setDate(new Date(fecha).getDate() + 1)).toLocaleDateString('es-ES');
  
    const saltadoresText = saltadores > 1 ? `${saltadores} SALTADORES` : `${saltadores} SALTADOR`;
    const kidsText = kids > 1 ? `${kids} KIDS` : `${kids} KID`;
    const noSaltadoresText = noSaltadores > 1 ? `${noSaltadores} ADULTOS` : `${noSaltadores} ADULTO`;
    const textoEExcel = [saltadoresText, kidsText, 'CUM', noSaltadoresText].join(' + ');
    const textoEClipboard = [
      saltadores > 1 ? `${saltadores} Saltadores` : `${saltadores} Saltador`,
      kids > 1 ? `${kids} Kids` : `${kids} Kid`,
      cumpleanero + ' (Bonificado)',
      noSaltadores > 1 ? `${noSaltadores} Adultos (${adultosExcluidos} Bonificados)` : `${noSaltadores} Adulto (${adultosExcluidos} Bonificados)`,
    ].filter(Boolean).join(' + ');
    const presupuestoTexto = `*${diaSemana} | ${formattedDate} | ${initialTime} A ${finalTime} | ${zona}*\n${textoEClipboard}\n*Presupuesto: $${presupuestoTotal.toLocaleString('es-ES')}*`;
  
    return { textoEExcel, presupuestoTexto, buffer: { nombre, cumpleanero, edad, telefono, textoEExcel, formattedDate, initialTime, finalTime, presupuestoTotal } };
  }
  
  document.getElementById('excelForm').addEventListener('submit', function (e) {
    e.preventDefault();
    
    const datos = obtenerDatosFormulario();
    if (!datos) return;
  
    const { nombre, cumpleanero, edad, telefono, saltadores, kids, noSaltadores, fecha, hora } = datos;
    const { textoEExcel, buffer } = calcularPresupuesto(datos);
  
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Datos');
  
    const defaultFont = { name: 'Calibri', size: 12 };
    const borderStyle = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
    const alignmentBottom = { vertical: 'bottom' };
    const alignmentCenter = { vertical: 'bottom', horizontal: 'center' };
    const confirmCellStyle = {
      font: { name: 'Calibri', size: 12, bold: true },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'C27BA0' }
      }
    };
  
    worksheet.addRow([nombre, cumpleanero, `${edad} años`, telefono, textoEExcel, 'A CONFIRMAR', buffer.formattedDate, `${buffer.initialTime} A ${buffer.finalTime}`, `$${buffer.presupuestoTotal.toLocaleString('es-ES')}`]);
  
    worksheet.eachRow((row, rowIndex) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        cell.font = defaultFont;
        cell.border = borderStyle;
        cell.alignment = alignmentBottom;
        if (colNumber >= 5) {
          cell.alignment = alignmentCenter;
        }
      });
    });
  
    const confirmCell = worksheet.getCell('F1');
    confirmCell.font = confirmCellStyle.font;
    confirmCell.fill = confirmCellStyle.fill;
  
    workbook.xlsx.writeBuffer().then(function (buffer) {
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'presupuesto.xlsx';
      a.click();
      window.URL.revokeObjectURL(url);
    });
  });
  
  document.getElementById('generateBudget').addEventListener('click', function () {
    const datos = obtenerDatosFormulario();
    if (!datos) return;
  
    const { presupuestoTexto } = calcularPresupuesto(datos);
  
    navigator.clipboard.writeText(presupuestoTexto).then(() => {
      alert('Presupuesto copiado al portapapeles.');
    }).catch(err => {
      alert('Error al copiar el presupuesto: ' + err);
    });
  });
  