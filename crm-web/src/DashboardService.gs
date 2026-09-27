/**
 * DashboardService.gs — Indicadores del Super Admin.
 */
const DashboardService = {
  summary(user) {
    const today = Utils.today();
    const month = today.slice(0, 7);
    const in7 = Utils.addDays(null, 7);
    const since30 = Utils.addDays(null, -30);

    const clientsAll = SearchIndex.records().filter(o => !o.FUSIONADO_EN);
    const clients = clientsAll.filter(o => Auth.canAccessClient(user, o));
    const allowed = {};
    clients.forEach(o => { allowed[o.CLIENTE_ID] = o; });
    const count = (list, field) => {
      const m = {};
      list.forEach(o => { const k = o[field] || '(vacío)'; m[k] = (m[k] || 0) + 1; });
      return m;
    };

    const opps = SheetService.readAll('OPORTUNIDADES').filter(o => allowed[o.CLIENTE_ID]);
    const open = opps.filter(o => o.ABIERTA);
    const porEtapa = {};
    ENUMS.ETAPA.forEach(e => {
      const l = opps.filter(o => o.ETAPA === e && (o.ABIERTA || String(o.FECHA_CIERRE_REAL).slice(0, 7) === month));
      porEtapa[e] = { label: STAGES[e].label, count: l.length, value: l.reduce((s, o) => s + Number(o.VALOR_ESTIMADO || 0), 0) };
    });
    const won = opps.filter(o => o.ETAPA === 'GANADO');
    const lost = opps.filter(o => o.ETAPA === 'PERDIDO' || o.ETAPA === 'NO_INTERESADO');

    const ventas = SheetService.readAll('VENTAS').filter(v => allowed[v.CLIENTE_ID] && v.ESTADO_PAGO !== 'ANULADA');
    const pagos = SheetService.readAll('PAGOS').filter(p => allowed[p.CLIENTE_ID]);
    const ventasMes = ventas.filter(v => String(v.FECHA_VENTA).slice(0, 7) === month);
    const months = [];
    for (let i = 5; i >= 0; i--) {
      const d = Utils.parseIsoDate(month + '-01');
      const m = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() - i, 15));
      months.push(Utilities.formatDate(m, 'UTC', 'yyyy-MM'));
    }
    const ventasPorMes = months.map(m => ({
      mes: m,
      ventas: ventas.filter(v => String(v.FECHA_VENTA).slice(0, 7) === m).reduce((s, v) => s + Number(v.VALOR_TOTAL || 0), 0),
      pagos: pagos.filter(p => String(p.FECHA_PAGO).slice(0, 7) === m).reduce((s, p) => s + Number(p.VALOR || 0), 0)
    }));

    const segs = SheetService.readColumns('SEGUIMIENTOS', ['SEGUIMIENTO_ID', 'CLIENTE_ID', 'TIPO', 'DESCRIPCION', 'FECHA_VENCIMIENTO', 'ESTADO', 'ASIGNADO_A', 'PRIORIDAD'])
      .filter(s => allowed[s.CLIENTE_ID] && s.ESTADO === 'PENDIENTE');
    const mine = segs.filter(s => s.ASIGNADO_A === user.email)
      .sort((a, b) => String(a.FECHA_VENCIMIENTO).localeCompare(String(b.FECHA_VENCIMIENTO)))
      .slice(0, 8)
      .map(s => ({
        SEGUIMIENTO_ID: s.SEGUIMIENTO_ID, CLIENTE_ID: s.CLIENTE_ID, TIPO: s.TIPO, DESCRIPCION: s.DESCRIPCION,
        FECHA_VENCIMIENTO: s.FECHA_VENCIMIENTO, PRIORIDAD: s.PRIORIDAD,
        CLIENTE_NOMBRE: allowed[s.CLIENTE_ID].NOMBRE_COMERCIAL || allowed[s.CLIENTE_ID].RAZON_SOCIAL
      }));

    const sols = SheetService.readColumns('SOLICITUDES', ['SOLICITUD_ID', 'RADICADO', 'CLIENTE_ID', 'CLIENTE_BASE_ID', 'ESTADO', 'ENVIADA_EN', 'INICIADA_EN', 'CAMBIOS_ESTADO'])
      .filter(s => {
        const cid = s.CLIENTE_ID || s.CLIENTE_BASE_ID;
        return !Auth.isScopedToOwn(user) || (cid && allowed[cid]);
      });
    const recientes = sols.filter(s => s.ESTADO !== 'BORRADOR')
      .sort((a, b) => String(b.ENVIADA_EN).localeCompare(String(a.ENVIADA_EN))).slice(0, 6)
      .map(s => ({
        SOLICITUD_ID: s.SOLICITUD_ID, RADICADO: s.RADICADO, CLIENTE_ID: s.CLIENTE_ID, ESTADO: s.ESTADO, ENVIADA_EN: s.ENVIADA_EN,
        CAMBIOS_ESTADO: s.CAMBIOS_ESTADO,
        CLIENTE_NOMBRE: allowed[s.CLIENTE_ID] ? (allowed[s.CLIENTE_ID].NOMBRE_COMERCIAL || allowed[s.CLIENTE_ID].RAZON_SOCIAL) : ''
      }));

    const acts = SheetService.readColumns('ACTIVIDADES', ['CLIENTE_ID', 'TIPO', 'FECHA_HORA'])
      .filter(a => allowed[a.CLIENTE_ID] && String(a.FECHA_HORA).slice(0, 10) >= since30);

    const propuestasEnviadas = {};
    SheetService.readColumns('PROPUESTAS', ['CLIENTE_ID', 'ESTADO']).forEach(p => {
      if (allowed[p.CLIENTE_ID] && (p.ESTADO === 'ENVIADA' || p.ESTADO === 'ACEPTADA')) propuestasEnviadas[p.CLIENTE_ID] = true;
    });

    return {
      clientes: {
        total: clients.length,
        porEstadoCrm: count(clients, 'ESTADO_CRM'),
        porMunicipio: count(clients, 'MUNICIPIO'),
        porEstadoFormulario: count(clients, 'ESTADO_FORMULARIO'),
        porOrigen: count(clients, 'ORIGEN'),
        sinEmail: clients.filter(o => !o.EMAIL).length,
        sinTelefono: clients.filter(o => !o.TELEFONO).length,
        noContactar: clients.filter(o => o.NO_CONTACTAR).length,
        sinResponsable: clients.filter(o => !o.RESPONSABLE_EMAIL).length
      },
      embudo: [
        { etapa: 'Prospectos', valor: clients.length },
        { etapa: 'Contactados', valor: clients.filter(o => o.ESTADO_CRM && ['NUEVO', 'POR_CONTACTAR'].indexOf(o.ESTADO_CRM) < 0).length },
        { etapa: 'Formularios enviados', valor: clients.filter(o => o.ESTADO_FORMULARIO && o.ESTADO_FORMULARIO !== 'NO_INICIADO').length },
        { etapa: 'Propuestas enviadas', valor: Object.keys(propuestasEnviadas).length },
        { etapa: 'Ventas', valor: ventas.length }
      ],
      oportunidades: {
        abiertas: open.length,
        valorPipeline: open.reduce((s, o) => s + Number(o.VALOR_ESTIMADO || 0), 0),
        valorPonderado: Math.round(open.reduce((s, o) => s + Number(o.VALOR_ESTIMADO || 0) * Number(o.PROBABILIDAD || 0) / 100, 0)),
        porEtapa: porEtapa,
        ganadasMes: won.filter(o => String(o.FECHA_CIERRE_REAL).slice(0, 7) === month).length,
        perdidasMes: lost.filter(o => String(o.FECHA_CIERRE_REAL).slice(0, 7) === month).length,
        tasaConversion: (won.length + lost.length) ? Math.round(100 * won.length / (won.length + lost.length)) : 0
      },
      ventas: {
        mesCantidad: ventasMes.length,
        mesValor: ventasMes.reduce((s, v) => s + Number(v.VALOR_TOTAL || 0), 0),
        pagosMes: pagos.filter(p => String(p.FECHA_PAGO).slice(0, 7) === month).reduce((s, p) => s + Number(p.VALOR || 0), 0),
        saldoPendiente: ventas.reduce((s, v) => s + Math.max(0, Number(v.VALOR_TOTAL || 0) - Number(v.VALOR_PAGADO || 0)), 0),
        porMes: ventasPorMes
      },
      seguimientos: {
        vencidos: segs.filter(s => s.FECHA_VENCIMIENTO < today).length,
        hoy: segs.filter(s => s.FECHA_VENCIMIENTO === today).length,
        proximos7: segs.filter(s => s.FECHA_VENCIMIENTO > today && s.FECHA_VENCIMIENTO <= in7).length,
        mios: mine
      },
      solicitudes: {
        recibidas30: sols.filter(s => s.ESTADO !== 'BORRADOR' && String(s.ENVIADA_EN).slice(0, 10) >= since30).length,
        porRevisar: sols.filter(s => s.ESTADO === 'ENVIADA').length,
        correccionesPendientes: sols.filter(s => s.CAMBIOS_ESTADO === 'PENDIENTE').length,
        borradoresActivos: sols.filter(s => s.ESTADO === 'BORRADOR' && String(s.INICIADA_EN).slice(0, 10) >= Utils.addDays(null, -7)).length,
        recientes: recientes
      },
      actividades30: count(acts, 'TIPO'),
      generado: Utils.nowIso()
    };
  }
};
