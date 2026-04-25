import pandas as pd
import numpy as np
from sqlalchemy import create_engine
import urllib
from datetime import datetime, timedelta
from pathlib import Path

import re
import sys

def _pedir_periodo():
    """Devuelve el período AAAA-MM: desde argv[1], o el mes actual si no se pasa argumento."""
    _default = datetime.now().strftime("%Y-%m")
    if len(sys.argv) > 1:
        _arg = sys.argv[1].strip()
        if re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", _arg):
            return _arg
        print(f"  Formato inválido '{_arg}'. Usa AAAA-MM, por ejemplo 2026-04.")
        sys.exit(1)
    return _default

PERIODO = _pedir_periodo()
print(f"Período: {PERIODO}")

# ══════════════════════════════════════════════════════════════
# 1. CARGA DE FUENTES
# ══════════════════════════════════════════════════════════════

# ── SQL Server: tabla principal ────────────────────────────────
params = urllib.parse.quote_plus(
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=AUREN22\\AUREN;"
    "DATABASE=eAuren;"
    "UID={_sql_user};"
    "PWD={_sql_password};"
)
engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}")

sql = """
DECLARE @periodo AS CHAR(7)
DECLARE @periodoAnterior AS CHAR(7)
SET @periodo = '2026-04';
SET @periodoAnterior = CONVERT(CHAR(7), DATEADD(MONTH, -1, CONVERT(DATE, @periodo + '-01')), 126);

WITH realme AS (
    SELECT 
        t.peticion,
        t.producto,
        t.sub_producto,
        t.segmento,
        t.zonal,
        'MASIVO' AS CANAL1,
        CASE 
            WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica FINANCIADO. (UP)%' THEN 'FLEX'
			WHEN t.[cms_desc_tipo_requerimiento] IS NULL THEN 'REGULAR'
			WHEN t.[cms_desc_tipo_requerimiento] = '0' THEN 'REGULAR'            
            ELSE 'REGULAR'
        END AS Scoring,
		t.vendedor AS VDD,
        t.documento_vendedor AS DNI_ORIG,
        CAST(t.fecha_registro AS DATE) AS Fecha_Registro,
        CAST(t.fecha_venta AS DATE) AS Fecha_Venta,
        CAST(a.fecha_alta AS DATE) AS Fecha_Alta,
        ROUND(t.pspaquete_renta_destino,1) AS Renta,
		t.usuario,
        t.tecnologia_ba,
        t.velocidad_ba,
        t.cms_codsrv,
		t.destinopaquete,
        cd.dni_vendedor AS dni_vendedor_cd,
        cx.dni_vendedor AS dni_vendedor_cx,
        COALESCE(cd.dni_vendedor, cx.dni_vendedor) AS dnivddcnet
    FROM fija_registros_totales t
        LEFT JOIN fija_altas a ON t.peticion = a.peticion
        LEFT JOIN fija_registros_unicos u ON t.peticion = u.peticion
        LEFT JOIN [dbo].[fija_controlnet_detallado] cd ON t.peticion = cd.nro_pedido
        LEFT JOIN [dbo].[fija_controlnet_detallado] cx ON t.cms_codsrv = cx.codigo_fe 
WHERE a.fecha_alta is NOT NULL AND t.categoria_producto = 'ALTA'
    AND (
        (FORMAT(t.fecha_registro, 'yyyy-MM') = @periodo AND FORMAT(a.fecha_alta, 'yyyy-MM') = @periodo)
        OR 
        (FORMAT(t.fecha_registro, 'yyyy-MM') = @periodoAnterior AND FORMAT(a.fecha_alta, 'yyyy-MM') = @periodo)
        OR
        (FORMAT(t.fecha_registro, 'yyyy-MM') = @periodo AND a.fecha_alta IS NULL)
    )
)
SELECT

Fecha_Alta AS Fecha_de_alta,
cms_codsrv AS FE,
peticion,
Q=1,
producto,
sub_producto,
CASE WHEN sub_producto IN ('TRIO','DUO INT+TV','COMPLETA TV') THEN 'TV' ELSE 'BA' END AS TV,
--COMBINACION
segmento,
NULL AS Geografia,
zonal,
usuario AS COM,
NULL AS DEPARTAMENTO,
NULL AS DOC_CLIENT,
NULL AS CLI_EXTRANJ,
NULL AS Canal,
NULL AS Entidad,
NULL AS Cod_Punto_Venta,
NULL AS Punto_Venta,
NULL AS RUC_Entidad,
NULL AS Razon_Social_Entidad,
CANAL1,
Scoring,
VDD,
DNI_ORIG,
NULL AS Zonal_Vendedor,
NULL AS Segmento_Cliente,
NULL AS Geografia_Cliente,
NULL AS Region_Cliente,
NULL AS Zonal_Cliente,
Fecha_Registro,
Fecha_Venta,
Fecha_Alta,
NULL AS Cod_PS_Origen,
NULL AS PS_Origen,
Renta,
NULL AS Cod_PS_Descuento_Ori,
NULL AS PS_Descuento_Ori,
NULL AS Renta_Descuento_Ori,
NULL AS LineaPS_Origen,
NULL AS LineaNombre_Origen,
NULL AS LineaRenta_Origen,
NULL AS SpeedyPS_Origen,
NULL AS SpeedyNombre_Origen,
NULL AS SpeedyRenta_Origen,
NULL AS CablePS_Origen,
NULL AS CableNombre_Origen,
NULL AS CableRenta_Origen,
NULL AS Desc_SVABa_Origen,
NULL AS Renta_SVABa_Origen,
NULL AS Cod_SVATv_Origen,
NULL AS Desc_SVATv_Origen,
NULL AS Renta_SVATv_Origen,
NULL AS decosventa_renta_origen,
NULL AS decosalquiler_renta_origen,
NULL AS decoscomodato_renta_origen,
NULL AS dsctopromocional_renta_origen,
NULL AS internetmovil_renta_origen,
NULL AS moden_renta_origen,
NULL AS multidestino_renta_origen,
NULL AS seguridadtotal_renta_origen,
NULL AS ultrawifi_renta_origen,
NULL AS arpuprincipal_origen,
NULL AS arpuxsvas_origen,
NULL AS RentaTotal_Origen,
NULL AS PSPaquete_Renta_Destino,
NULL AS Cod_PS_Descuento_Des,
NULL AS PS_Descuento_Des,
NULL AS Renta_Descuento_Des,
NULL AS LineaPS_Destino,
NULL AS LineaNombre_Destino,
NULL AS LineaRenta_Destino,
NULL AS SpeedyPS_Destino,
NULL AS SpeedyNombre_Destino,
NULL AS SpeedyRenta_Destino,
NULL AS CablePS_Destino,
NULL AS CableNombre_Destino,
NULL AS CableRenta_Destino,
NULL AS Desc_SVABa_Destino,
NULL AS Renta_SVABa_Destino,
NULL AS Cod_SVATv_Destino,
NULL AS Desc_SVATv_Destino,
NULL AS Renta_SVATv_Destino,
NULL AS decosventa_renta_destino,
NULL AS decosalquiler_renta_destino,
NULL AS decoscomodato_renta_destino,
NULL AS dsctopromocional_renta_destino,
NULL AS internetmovil_renta_destino,
NULL AS moden_renta_destino,
NULL AS multidestino_renta_destino,
NULL AS seguridadtotal_renta_destino,
NULL AS ultrawifi_renta_destino,
NULL AS arpuprincipal_destino,
NULL AS arpuxsvas_destino,
NULL AS saltoxsvas,
NULL AS saltototal,
NULL AS RentaTotal_Destino,
NULL AS Ingresos,
NULL AS Cluster_Origen,
NULL AS Cluster_Destino,
NULL AS Componentes_Alta,
NULL AS GAP_ARPU,
NULL AS Tipo_GAP,
NULL AS Source_System,
NULL AS Flag_Duplicado,
NULL AS Flag_Registro,
NULL AS Flag_Registro_Unico,
NULL AS Flag_Venta,
NULL AS Flag_Alta,
NULL AS Flag_Web,
NULL AS Producto_Web,
NULL AS Canal_Web,
NULL AS Flag_Web_Alta,
NULL AS RU_Primera_Peticion,
NULL AS RU_Primera_Fecha_Registro,
NULL AS RU_Primer_Canal,
NULL AS Tecnologia_TV,
tecnologia_ba,
velocidad_ba,
NULL AS Marca_Back,
NULL AS Telefono_contacto_1,
NULL AS Telefono_contacto_2,
NULL AS PETICIONES_PRUEBA,
NULL AS ID_CLIENTE,
NULL AS CUENTA,
NULL AS NUMERO_PETICION,
NULL AS ABONADO,
NULL AS INSCRIPCION,
NULL AS CUENTAFACTURACIONCD,
NULL AS SEGMENTOCD,
NULL AS MACROSEGMENTO,
NULL AS CICLO,
NULL AS DIRECCIONCD,
NULL AS SUBLOCALIDADCD,
NULL AS PROVINCIACD,
NULL AS MUNICIPALIDADCD,
NULL AS DOCUMENTO,
NULL AS RUC,
NULL AS NUMERO_TELEFONO,
NULL AS FEC_REG_PETICION,
NULL AS FECHACAMBIOESTADO,
NULL AS ESTADO_PETICION,
NULL AS DESC_ESTADO_PETICION,
NULL AS ESTADO_SUBPET_ORIGEN,
NULL AS ESTADO_SUBPET_DESTINO,
NULL AS DESC_ESTADO_SUBPET_DESTINO,
NULL AS ESTADO_AGRUP_ORIGEN,
NULL AS ESTADO_AGRUP_DESTINO,
NULL AS DESC_ESTADO_AGRUP_DESTINO,
NULL AS MOTIVO_ESTADO,
NULL AS DESC_MOTIVO_ESTADO,
NULL AS MOTIVO_CANCELACION,
NULL AS DESC_MOTIVO_CANCELACION,
NULL AS SUBMOTIVO_CANCELACION,
NULL AS DESC_SUBMOTIVO_CANCELACION,
NULL AS ID_CANAL_VENTA,
NULL AS ID_PUNTO_VENTA,
NULL AS ID_VENDEDOR,
NULL AS USUARIO2,
NULL AS TIP_OPE_CMR_SUBPET_ORIGEN,
NULL AS TIP_OPE_CMR_SUBPET_DESTINO,
NULL AS TIPO_USO_ORIGEN,
NULL AS TIPO_USO_DESTINO,
NULL AS FLAG_BAJA_DEUDA,
NULL AS LOCALIDAD,
NULL AS AREA,
NULL AS ORIGENPAQUETE,
destinopaquete AS DESTINOPAQUETE,
NULL AS ORIGENLINEA,
NULL AS DESTINOLINEA,
NULL AS ORIGENINTERNET,
NULL AS DESTINOINTERNET,
NULL AS ORIGENCABLE,
NULL AS DESTINOCABLE,
NULL AS ORIGENBLOQUESTV,
NULL AS DESTINOBLOQUESTV,
NULL AS ORIGENDESCUENTO,
NULL AS DESTINODESCUENTO,
NULL AS ORIGENBLOQUESTVDS,
NULL AS DESTINOBLOQUESTVDS,
NULL AS ORIGENMODEM,
NULL AS DESTINOMODEM,
NULL AS ORIGENDECOVENTA,
NULL AS DESTINODECOVENTA,
NULL AS ORIGENDECOALQUILER,
NULL AS DESTINODECOALQUILER,
NULL AS ORIGENDECOCOMODATO,
NULL AS DESTINODECOCOMODATO,
NULL AS ORIGENUWIFI,
NULL AS DESTINOUWIFI,
NULL AS ORIGENDESCUENTOTEMPORAL,
NULL AS DESTINODESCUENTOTEMPORAL,
NULL AS ORIGENINTERNETMOVIL,
NULL AS DESTINOINTERNETMOVIL,
NULL AS ORIGENMULTIDESTINO,
NULL AS DESTINOMULTIDESTINO,
NULL AS ORIGENSEGURIDADTOTAL,
NULL AS DESTINOSEGURIDADTOTAL,
NULL AS ORIGENPSADMINISTRATIVA,
NULL AS DESTINOPSADMINISTRATIVA,
NULL AS ORIGENCLUSTER,
NULL AS DESTINOCLUSTER,
NULL AS ORIGENTECNOLOGIAINTERNET,
NULL AS DESTINOTECNOLOGIAINTERNET,
NULL AS ORIGENTECNOLOGIATV,
NULL AS DESTINOTECNOLOGIATV,
NULL AS ORIGENDECOSIGLAS,
NULL AS DESTINODECOSIGLAS,
NULL AS REQUE,
NULL AS ORIGEN_REQUE,
NULL AS TIPO_REQUE,
NULL AS MOTIVO_REQUE,
NULL AS ESTADO_REQUE,
NULL AS OFICINA_ADMINISTRATIVA,
NULL AS SITUACION_REQUE,
NULL AS CONDICION_REQUE,
NULL AS FEC_REGISTRO,
NULL AS FEC_LIQUIDACION,
NULL AS CLIENTE_CMS,
NULL AS CUENTA_CMS,
NULL AS SERVICIO_CMS,
NULL AS CLASE_SERVICIO,
NULL AS OFERTA_CMS,
NULL AS USUARIO_CMS,
NULL AS TELEFONO_VOIP,
NULL AS TIPO_PAQUETE_MULTI,
NULL AS FECHAFORMULACION,
NULL AS SEGMENTO_CUENTA,
NULL AS CMS_CODREQ,
NULL AS CMS_INDORIGREQ,
NULL AS CMS_TIPREQ,
cms_codsrv,
NULL AS CMS_CODCLI,
NULL AS CMS_FECREG,
NULL AS CMS_FECLIQ,
NULL AS CMS_FECEST,
NULL AS CMS_FECASG,
NULL AS CMS_FECPRG,
NULL AS CMS_CODEDO,
NULL AS CMS_CODSIT,
NULL AS CMS_CODMOTV,
NULL AS CMS_CODUSR,
NULL AS CMS_CODOFEPROD,
NULL AS CMS_PLAZOCNTR,
NULL AS CMS_DESTIPREQ,
NULL AS CMS_CODGRPREQ,
NULL AS CMS_DESGRPREQ,
NULL AS CMS_COD_PETICION,
NULL AS CMS_NUMTELEFVOIP,
NULL AS CMS_TIPPAQMUL,
NULL AS CMS_DISTRITOCD,
NULL AS CMS_CODNOD,
NULL AS CMS_NROPLANO,
NULL AS CMS_CODCLASRV,
NULL AS CMS_DESCLASRV,
NULL AS CMS_DESTINOTV,
NULL AS CMS_DESTINOBLOQUE,
NULL AS CMS_DESTINOINTERNET,
NULL AS CMS_DESTINOVOZIP,
NULL AS CMS_ORIGENTV,
NULL AS CMS_ORIGENBLOQUE,
NULL AS CMS_ORIGENINTERNET,
NULL AS CMS_ORIGENVOZIP,
NULL AS CMS_DESTINOVELOCIDADINTERNET,
NULL AS CMS_ORIGENVELOCIDADINTERNET,
NULL AS CMS_DESTINODECOALQUILER,
NULL AS CMS_DESTINODECOVENTA,
NULL AS CMS_DESTINODECOCOMODATO,
NULL AS CMS_DESTINODECOOTROS,
NULL AS CMS_ORIGENDECOALQUILER,
NULL AS CMS_ORIGENDECOVENTA,
NULL AS CMS_ORIGENDECOCOMODATO,
NULL AS CMS_ORIGENDECOOTROS,
NULL AS CMS_ORIGENMODEM,
NULL AS CMS_DESTINOMODEM,
NULL AS CMS_DESTINOMACADDRESSDECOALQUILER,
NULL AS CMS_DESTINOMACADDRESSDECOVENTA,
NULL AS CMS_DESTINOMACADDRESSDECOCOMODATO,
NULL AS CMS_DESTINOMACADDRESSDECOOTROS,
NULL AS CMS_ORIGENMACADDRESSDECOALQUILER,
NULL AS CMS_ORIGENMACADDRESSDECOVENTA,
NULL AS CMS_ORIGENMACADDRESSDECOCOMODATO,
NULL AS CMS_ORIGENMACADDRESSDECOOTROS,
NULL AS CMS_ORIGENMACADDRESSMODEM,
NULL AS CMS_DESTINMACADDRESSOMODEM,
NULL AS CMS_CODCNLVTA,
NULL AS CMS_CODGRPVTA,
NULL AS CMS_CODVDD,
NULL AS CMS_CODMOTACT,
NULL AS CMS_CODUSRGEN,
NULL AS CMS_CODCTR,
NULL AS CMS_DOCUMENTO,
NULL AS CMS_TIPODOC,
NULL AS CMS_RUC,
NULL AS CMS_SEGMENTO,
NULL AS CMS_TELCONTACTO,
NULL AS MACROSEGMENTOTOP,
NULL AS GESTIONDECOSTOTAL,
NULL AS GESTIONBLOQUESTV,
NULL AS GESTIONULTRAWIFI,
NULL AS TIPOTRANSACCION,
NULL AS CANALAGRUPADO,
NULL AS CMS_DESC_TIPO_REQUERIMIENTO,
NULL AS CMS_DESC_SITUACION_REQUERIMIENTO,
NULL AS CMS_DESC_MOTIVO_GENERACION,
NULL AS CMS_DESC_PRODUCTO,
NULL AS Region_Vendedor2,
NULL AS MATCH_CONVER,
NULL AS TRUEFALSE,
'S' + CAST(DATEPART(WEEK, fecha_alta) - DATEPART(WEEK, DATEADD(MONTH, DATEDIFF(MONTH, 0, fecha_alta), 0)) + 1 AS VARCHAR(2)) AS Semana,
dnivddcnet
FROM realme
ORDER BY Fecha_Registro ASC;
"""

sql_rt = """
DECLARE @periodo AS CHAR(7)
DECLARE @periodoAnterior AS CHAR(7)
SET @periodo = '2026-04';
SET @periodoAnterior = CONVERT(CHAR(7), DATEADD(MONTH, -1, CONVERT(DATE, @periodo + '-01')), 126);

WITH realme AS (
    SELECT
        t.peticion,
        t.producto,
        t.sub_producto,
        t.segmento,
        t.zonal,
        'MASIVO' AS CANAL1,
        CASE
            WHEN SUBSTRING(t.[cms_desc_tipo_requerimiento],CHARINDEX('Fija',t.[cms_desc_tipo_requerimiento]),29) LIKE 'Fija, aplica FINANCIADO. (UP)%' THEN 'FLEX'
			WHEN t.[cms_desc_tipo_requerimiento] IS NULL THEN 'REGULAR'
			WHEN t.[cms_desc_tipo_requerimiento] = '0' THEN 'REGULAR'
            ELSE 'REGULAR'
        END AS Scoring,
		t.vendedor AS VDD,
        t.documento_vendedor AS DNI_ORIG,
        CAST(t.fecha_registro AS DATE) AS Fecha_Registro,
        CAST(t.fecha_venta AS DATE) AS Fecha_Venta,
        CAST(t.fecha_alta AS DATE) AS Fecha_Alta,
        ROUND(t.pspaquete_renta_destino,1) AS Renta,
		t.usuario,
        t.tecnologia_ba,
        t.velocidad_ba,
        t.cms_codsrv,
		t.destinopaquete,
        cd.dni_vendedor AS dni_vendedor_cd,
        cx.dni_vendedor AS dni_vendedor_cx,
        COALESCE(cd.dni_vendedor, cx.dni_vendedor) AS dnivddcnet
    FROM fija_registros_totales t
        LEFT JOIN [dbo].[fija_controlnet_detallado] cd ON t.peticion = cd.nro_pedido
        LEFT JOIN [dbo].[fija_controlnet_detallado] cx ON t.cms_codsrv = cx.codigo_fe
WHERE t.categoria_producto = 'ALTA' AND FORMAT(t.fecha_registro, 'yyyy-MM') = @periodo)

SELECT

Fecha_Alta AS Fecha_de_alta,
cms_codsrv AS FE,
peticion,
Q=1,
producto,
sub_producto,
CASE WHEN sub_producto IN ('TRIO','DUO INT+TV','COMPLETA TV') THEN 'TV' ELSE 'BA' END AS TV,
--COMBINACION
segmento,
NULL AS Geografia,
NULL AS RIESG,
zonal,
usuario AS COM,
NULL AS DEPARTAMENTO,
NULL AS DOC_CLIENT,
NULL AS CLI_EXTRANJ,
NULL AS Canal,
NULL AS Entidad,
NULL AS Cod_Punto_Venta,
NULL AS Punto_Venta,
NULL AS RUC_Entidad,
NULL AS Razon_Social_Entidad,
CANAL1,
Scoring,
VDD,
DNI_ORIG,
NULL AS Zonal_Vendedor,
NULL AS Segmento_Cliente,
NULL AS Geografia_Cliente,
NULL AS Region_Cliente,
NULL AS Zonal_Cliente,
Fecha_Registro,
Fecha_Venta,
Fecha_Alta,
NULL AS Cod_PS_Origen,
NULL AS PS_Origen,
Renta,
NULL AS Cod_PS_Descuento_Ori,
NULL AS PS_Descuento_Ori,
NULL AS Renta_Descuento_Ori,
NULL AS LineaPS_Origen,
NULL AS LineaNombre_Origen,
NULL AS LineaRenta_Origen,
NULL AS SpeedyPS_Origen,
NULL AS SpeedyNombre_Origen,
NULL AS SpeedyRenta_Origen,
NULL AS CablePS_Origen,
NULL AS CableNombre_Origen,
NULL AS CableRenta_Origen,
NULL AS Desc_SVABa_Origen,
NULL AS Renta_SVABa_Origen,
NULL AS Cod_SVATv_Origen,
NULL AS Desc_SVATv_Origen,
NULL AS Renta_SVATv_Origen,
NULL AS decosventa_renta_origen,
NULL AS decosalquiler_renta_origen,
NULL AS decoscomodato_renta_origen,
NULL AS dsctopromocional_renta_origen,
NULL AS internetmovil_renta_origen,
NULL AS moden_renta_origen,
NULL AS multidestino_renta_origen,
NULL AS seguridadtotal_renta_origen,
NULL AS ultrawifi_renta_origen,
NULL AS arpuprincipal_origen,
NULL AS arpuxsvas_origen,
NULL AS RentaTotal_Origen,
NULL AS PSPaquete_Renta_Destino,
NULL AS Cod_PS_Descuento_Des,
NULL AS PS_Descuento_Des,
NULL AS Renta_Descuento_Des,
NULL AS LineaPS_Destino,
NULL AS LineaNombre_Destino,
NULL AS LineaRenta_Destino,
NULL AS SpeedyPS_Destino,
NULL AS SpeedyNombre_Destino,
NULL AS SpeedyRenta_Destino,
NULL AS CablePS_Destino,
NULL AS CableNombre_Destino,
NULL AS CableRenta_Destino,
NULL AS Desc_SVABa_Destino,
NULL AS Renta_SVABa_Destino,
NULL AS Cod_SVATv_Destino,
NULL AS Desc_SVATv_Destino,
NULL AS Renta_SVATv_Destino,
NULL AS decosventa_renta_destino,
NULL AS decosalquiler_renta_destino,
NULL AS decoscomodato_renta_destino,
NULL AS dsctopromocional_renta_destino,
NULL AS internetmovil_renta_destino,
NULL AS moden_renta_destino,
NULL AS multidestino_renta_destino,
NULL AS seguridadtotal_renta_destino,
NULL AS ultrawifi_renta_destino,
NULL AS arpuprincipal_destino,
NULL AS arpuxsvas_destino,
NULL AS saltoxsvas,
NULL AS saltototal,
NULL AS RentaTotal_Destino,
NULL AS Ingresos,
NULL AS Cluster_Origen,
NULL AS Cluster_Destino,
NULL AS Componentes_Alta,
NULL AS GAP_ARPU,
NULL AS Tipo_GAP,
NULL AS Source_System,
NULL AS Flag_Duplicado,
NULL AS Flag_Registro,
NULL AS Flag_Registro_Unico,
NULL AS Flag_Venta,
NULL AS Flag_Alta,
NULL AS Flag_Web,
NULL AS Producto_Web,
NULL AS Canal_Web,
NULL AS Flag_Web_Alta,
NULL AS RU_Primera_Peticion,
NULL AS RU_Primera_Fecha_Registro,
NULL AS RU_Primer_Canal,
NULL AS Tecnologia_TV,
tecnologia_ba,
velocidad_ba,
NULL AS Marca_Back,
NULL AS Telefono_contacto_1,
NULL AS Telefono_contacto_2,
NULL AS PETICIONES_PRUEBA,
NULL AS ID_CLIENTE,
NULL AS CUENTA,
NULL AS NUMERO_PETICION,
NULL AS ABONADO,
NULL AS INSCRIPCION,
NULL AS CUENTAFACTURACIONCD,
NULL AS SEGMENTOCD,
NULL AS MACROSEGMENTO,
NULL AS CICLO,
NULL AS DIRECCIONCD,
NULL AS SUBLOCALIDADCD,
NULL AS PROVINCIACD,
NULL AS MUNICIPALIDADCD,
NULL AS DOCUMENTO,
NULL AS RUC,
NULL AS NUMERO_TELEFONO,
NULL AS FEC_REG_PETICION,
NULL AS FECHACAMBIOESTADO,
NULL AS ESTADO_PETICION,
NULL AS DESC_ESTADO_PETICION,
NULL AS ESTADO_SUBPET_ORIGEN,
NULL AS ESTADO_SUBPET_DESTINO,
NULL AS DESC_ESTADO_SUBPET_DESTINO,
NULL AS ESTADO_AGRUP_ORIGEN,
NULL AS ESTADO_AGRUP_DESTINO,
NULL AS DESC_ESTADO_AGRUP_DESTINO,
NULL AS MOTIVO_ESTADO,
NULL AS DESC_MOTIVO_ESTADO,
NULL AS MOTIVO_CANCELACION,
NULL AS DESC_MOTIVO_CANCELACION,
NULL AS SUBMOTIVO_CANCELACION,
NULL AS DESC_SUBMOTIVO_CANCELACION,
NULL AS ID_CANAL_VENTA,
NULL AS ID_PUNTO_VENTA,
NULL AS ID_VENDEDOR,
NULL AS USUARIO2,
NULL AS TIP_OPE_CMR_SUBPET_ORIGEN,
NULL AS TIP_OPE_CMR_SUBPET_DESTINO,
NULL AS TIPO_USO_ORIGEN,
NULL AS TIPO_USO_DESTINO,
NULL AS FLAG_BAJA_DEUDA,
NULL AS LOCALIDAD,
NULL AS AREA,
NULL AS ORIGENPAQUETE,
destinopaquete AS DESTINOPAQUETE,
NULL AS ORIGENLINEA,
NULL AS DESTINOLINEA,
NULL AS ORIGENINTERNET,
NULL AS DESTINOINTERNET,
NULL AS ORIGENCABLE,
NULL AS DESTINOCABLE,
NULL AS ORIGENBLOQUESTV,
NULL AS DESTINOBLOQUESTV,
NULL AS ORIGENDESCUENTO,
NULL AS DESTINODESCUENTO,
NULL AS ORIGENBLOQUESTVDS,
NULL AS DESTINOBLOQUESTVDS,
NULL AS ORIGENMODEM,
NULL AS DESTINOMODEM,
NULL AS ORIGENDECOVENTA,
NULL AS DESTINODECOVENTA,
NULL AS ORIGENDECOALQUILER,
NULL AS DESTINODECOALQUILER,
NULL AS ORIGENDECOCOMODATO,
NULL AS DESTINODECOCOMODATO,
NULL AS ORIGENUWIFI,
NULL AS DESTINOUWIFI,
NULL AS ORIGENDESCUENTOTEMPORAL,
NULL AS DESTINODESCUENTOTEMPORAL,
NULL AS ORIGENINTERNETMOVIL,
NULL AS DESTINOINTERNETMOVIL,
NULL AS ORIGENMULTIDESTINO,
NULL AS DESTINOMULTIDESTINO,
NULL AS ORIGENSEGURIDADTOTAL,
NULL AS DESTINOSEGURIDADTOTAL,
NULL AS ORIGENPSADMINISTRATIVA,
NULL AS DESTINOPSADMINISTRATIVA,
NULL AS ORIGENCLUSTER,
NULL AS DESTINOCLUSTER,
NULL AS ORIGENTECNOLOGIAINTERNET,
NULL AS DESTINOTECNOLOGIAINTERNET,
NULL AS ORIGENTECNOLOGIATV,
NULL AS DESTINOTECNOLOGIATV,
NULL AS ORIGENDECOSIGLAS,
NULL AS DESTINODECOSIGLAS,
NULL AS REQUE,
NULL AS ORIGEN_REQUE,
NULL AS TIPO_REQUE,
NULL AS MOTIVO_REQUE,
NULL AS ESTADO_REQUE,
NULL AS OFICINA_ADMINISTRATIVA,
NULL AS SITUACION_REQUE,
NULL AS CONDICION_REQUE,
NULL AS FEC_REGISTRO,
NULL AS FEC_LIQUIDACION,
NULL AS CLIENTE_CMS,
NULL AS CUENTA_CMS,
NULL AS SERVICIO_CMS,
NULL AS CLASE_SERVICIO,
NULL AS OFERTA_CMS,
NULL AS USUARIO_CMS,
NULL AS TELEFONO_VOIP,
NULL AS TIPO_PAQUETE_MULTI,
NULL AS FECHAFORMULACION,
NULL AS SEGMENTO_CUENTA,
NULL AS CMS_CODREQ,
NULL AS CMS_INDORIGREQ,
NULL AS CMS_TIPREQ,
cms_codsrv,
NULL AS CMS_CODCLI,
NULL AS CMS_FECREG,
NULL AS CMS_FECLIQ,
NULL AS CMS_FECEST,
NULL AS CMS_FECASG,
NULL AS CMS_FECPRG,
NULL AS CMS_CODEDO,
NULL AS CMS_CODSIT,
NULL AS CMS_CODMOTV,
NULL AS CMS_CODUSR,
NULL AS CMS_CODOFEPROD,
NULL AS CMS_PLAZOCNTR,
NULL AS CMS_DESTIPREQ,
NULL AS CMS_CODGRPREQ,
NULL AS CMS_DESGRPREQ,
NULL AS CMS_COD_PETICION,
NULL AS CMS_NUMTELEFVOIP,
NULL AS CMS_TIPPAQMUL,
NULL AS CMS_DISTRITOCD,
NULL AS CMS_CODNOD,
NULL AS CMS_NROPLANO,
NULL AS CMS_CODCLASRV,
NULL AS CMS_DESCLASRV,
NULL AS CMS_DESTINOTV,
NULL AS CMS_DESTINOBLOQUE,
NULL AS CMS_DESTINOINTERNET,
NULL AS CMS_DESTINOVOZIP,
NULL AS CMS_ORIGENTV,
NULL AS CMS_ORIGENBLOQUE,
NULL AS CMS_ORIGENINTERNET,
NULL AS CMS_ORIGENVOZIP,
NULL AS CMS_DESTINOVELOCIDADINTERNET,
NULL AS CMS_ORIGENVELOCIDADINTERNET,
NULL AS CMS_DESTINODECOALQUILER,
NULL AS CMS_DESTINODECOVENTA,
NULL AS CMS_DESTINODECOCOMODATO,
NULL AS CMS_DESTINODECOOTROS,
NULL AS CMS_ORIGENDECOALQUILER,
NULL AS CMS_ORIGENDECOVENTA,
NULL AS CMS_ORIGENDECOCOMODATO,
NULL AS CMS_ORIGENDECOOTROS,
NULL AS CMS_ORIGENMODEM,
NULL AS CMS_DESTINOMODEM,
NULL AS CMS_DESTINOMACADDRESSDECOALQUILER,
NULL AS CMS_DESTINOMACADDRESSDECOVENTA,
NULL AS CMS_DESTINOMACADDRESSDECOCOMODATO,
NULL AS CMS_DESTINOMACADDRESSDECOOTROS,
NULL AS CMS_ORIGENMACADDRESSDECOALQUILER,
NULL AS CMS_ORIGENMACADDRESSDECOVENTA,
NULL AS CMS_ORIGENMACADDRESSDECOCOMODATO,
NULL AS CMS_ORIGENMACADDRESSDECOOTROS,
NULL AS CMS_ORIGENMACADDRESSMODEM,
NULL AS CMS_DESTINMACADDRESSOMODEM,
NULL AS CMS_CODCNLVTA,
NULL AS CMS_CODGRPVTA,
NULL AS CMS_CODVDD,
NULL AS CMS_CODMOTACT,
NULL AS CMS_CODUSRGEN,
NULL AS CMS_CODCTR,
NULL AS CMS_DOCUMENTO,
NULL AS CMS_TIPODOC,
NULL AS CMS_RUC,
NULL AS CMS_SEGMENTO,
NULL AS CMS_TELCONTACTO,
NULL AS MACROSEGMENTOTOP,
NULL AS GESTIONDECOSTOTAL,
NULL AS GESTIONBLOQUESTV,
NULL AS GESTIONULTRAWIFI,
NULL AS TIPOTRANSACCION,
NULL AS CANALAGRUPADO,
NULL AS CMS_DESC_TIPO_REQUERIMIENTO,
NULL AS CMS_DESC_SITUACION_REQUERIMIENTO,
NULL AS CMS_DESC_MOTIVO_GENERACION,
NULL AS CMS_DESC_PRODUCTO,
NULL AS Region_Vendedor2,
NULL AS MATCH_CONVER,
NULL AS TRUEFALSE,
NULL AS Semana,
dnivddcnet
FROM realme
ORDER BY Fecha_Registro ASC;
"""

sql_con = f"""
SELECT
    [tipo],
    CAST([documento_v] AS INT) AS DNIVDD,
    SUM(CAST([consulta_unica] AS INT)) AS Q
FROM [eAuren].[dbo].[fija_base_dito_consultas_hoy]
WHERE [periodo] = '{PERIODO}' AND TIPO = 'INTENCIONES'
GROUP BY [tipo], [documento_v], [consulta_unica]
"""

PERIODO_ANT = (
    pd.Timestamp(PERIODO + "-01") - pd.offsets.MonthBegin(1)
).strftime("%Y-%m")

sql    = sql.replace(   "SET @periodo = '2026-04';", f"SET @periodo = '{PERIODO}';")
sql_rt = sql_rt.replace("SET @periodo = '2026-04';", f"SET @periodo = '{PERIODO}';")

# SQL mes anterior: misma query principal con PERIODO_ANT
sql_ant = sql.replace(f"SET @periodo = '{PERIODO}';", f"SET @periodo = '{PERIODO_ANT}';")

df        = pd.read_sql(sql,     engine)
df_rt     = pd.read_sql(sql_rt,  engine)
df_con    = pd.read_sql(sql_con, engine)
df_ant    = pd.read_sql(sql_ant, engine)   # altas del mes anterior
engine.dispose()

# ── Google Sheets: VENTORY y RH ────────────────────────────────
URL_VENTORY = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSXxqrGGs4_mU4n511v3zBkKo4buAFv0TwrlrrX4XD2jFjIT7cC8kvH7ER32Ye2hiOpo3mAFsUkyydg/pub?gid=22270598&single=true&output=csv"
URL_RH      = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSXxqrGGs4_mU4n511v3zBkKo4buAFv0TwrlrrX4XD2jFjIT7cC8kvH7ER32Ye2hiOpo3mAFsUkyydg/pub?gid=241856834&single=true&output=csv"
URL_MF      = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQoJkR0gbsSLkLd2vv1JjtARboRLdaKLM64MPnuxyyZmLbAOEcA6jCuz9hBV7FfRT1qSajTxaEYyOhs/pub?gid=1358673506&single=true&output=csv"

ventory = pd.read_csv(URL_VENTORY)
rh      = pd.read_csv(URL_RH)
mf      = pd.read_csv(URL_MF)

hist = pd.read_excel(Path(__file__).parent / "altas_historico.xlsx")
hist["dnivdd"] = pd.to_numeric(hist["dnivdd"], errors="coerce").astype("Int64")

# ══════════════════════════════════════════════════════════════
# 2. LIMPIEZA DE FUENTES
# ══════════════════════════════════════════════════════════════

# ── Tabla principal ────────────────────────────────────────────
df = df.drop_duplicates(subset="peticion")

date_cols = ["Fecha_de_alta", "Fecha_Registro", "Fecha_Venta", "Fecha_Alta"]
for c in date_cols:
    df[c] = pd.to_datetime(df[c], errors="coerce")

df["DNI_ORIG"] = pd.to_numeric(df["DNI_ORIG"], errors="coerce").astype("Int64")
df["peticion"] = pd.to_numeric(df["peticion"], errors="coerce").astype("Int64")
df["FE"]       = df["FE"].astype(str).str.strip()

# ── VENTORY ────────────────────────────────────────────────────
ventory["Fecha_Registro"] = pd.to_datetime(
    ventory["Fecha_Registro"], dayfirst=True, errors="coerce"
)
ventory["dnivdd"]    = pd.to_numeric(ventory["dnivdd"],    errors="coerce").astype("Int64")
ventory["peticion"]  = pd.to_numeric(ventory["peticion"],  errors="coerce").astype("Int64")
ventory["codigo_fe"] = ventory["codigo_fe"].astype(str).str.strip()

# ── RH ─────────────────────────────────────────────────────────
rh = rh.drop(columns=["JEFE", "F_ING"], errors="ignore")
rh["DNI"] = pd.to_numeric(rh["DNI"], errors="coerce").astype("Int64")
if "SUPERVISOR" in rh.columns:
    rh["SUPERVISOR"] = (rh["SUPERVISOR"]
                        .astype(str)
                        .str.upper()
                        .str.replace(r"\s{2,}", " ", regex=True)
                        .str.strip()
                        .str.translate(str.maketrans("ÁÉÍÓÚ", "AEIOU")))

# ══════════════════════════════════════════════════════════════
# 3. JOIN DOBLE CON VENTORY
#    Prioriza el registro más reciente en caso de duplicados
#    Primero por codigo_fe = FE, luego por peticion como fallback
# ══════════════════════════════════════════════════════════════

# Join 1: por codigo_fe → FE (registro más reciente por FE)
ventory_por_fe = (
    ventory
    .sort_values("Fecha_Registro", ascending=False)
    .drop_duplicates(subset="codigo_fe")
    [["codigo_fe", "dnivdd"]]
    .rename(columns={"codigo_fe": "FE", "dnivdd": "v1_dnivdd"})
)

# Join 2: por peticion (registro más reciente por peticion)
ventory_por_pet = (
    ventory
    .sort_values("Fecha_Registro", ascending=False)
    .drop_duplicates(subset="peticion")
    [["peticion", "dnivdd"]]
    .rename(columns={"dnivdd": "v2_dnivdd"})
)

df = df.merge(ventory_por_fe, on="FE", how="left")
df = df.merge(ventory_por_pet, on="peticion", how="left")

# Coalesce: FE tiene prioridad, peticion es fallback
df["dnivdd_final"] = df["v1_dnivdd"].combine_first(df["v2_dnivdd"])

# Flag CTRLNET
df["CTRLNET"] = df["dnivdd_final"].notna().map({True: "SI", False: "NO"})

# Limpieza columnas temporales
df = df.drop(columns=["v1_dnivdd", "v2_dnivdd"])

# ══════════════════════════════════════════════════════════════
# 4. JOIN CON RH
#    dnivdd_final = DNI → trae ZONAL, SUPERVISOR, ESQUEMA, VENDEDOR
# ══════════════════════════════════════════════════════════════

rh_slim = (
    rh[["DNI", "ZONA", "SUPERVISOR", "ESQUEMA", "VENDEDOR"]]  # ZONA, no ZONAL
    .drop_duplicates(subset="DNI")
    .rename(columns={
        "DNI":        "dnivdd_final",
        "ZONA":       "RH.ZONAL",      # aquí le ponemos el nombre final
        "SUPERVISOR": "RH.SUPERVISOR",
        "ESQUEMA":    "RH.ESQUEMA",
        "VENDEDOR":   "RH.VENDEDOR",
    })
)

df = df.merge(rh_slim, on="dnivdd_final", how="left")

# ══════════════════════════════════════════════════════════════
# 5. COLUMNAS CALCULADAS
# ══════════════════════════════════════════════════════════════

tv_productos = {"TRIO", "DUO INT+TV", "COMPLETA TV"}

df["DNI_VENDEDOR"]    = df["dnivdd_final"].combine_first(df["DNI_ORIG"]).astype("Int64")
df["PilotoPR+MONOBA"] = (df["Scoring"] == "FLEX").astype(int)
df["TV"]              = df["sub_producto"].apply(
                            lambda x: "TV" if x in tv_productos else "BA"
                        )
df["MATCH_DIRECC"]    = df.apply(
    lambda r: None if pd.isna(r["dnivdd_final"])
              else (1 if r["dnivdd_final"] != r["DNI_ORIG"] else 0),
    axis=1
)

# ══════════════════════════════════════════════════════════════
# 6. RENOMBRES FINALES
# ══════════════════════════════════════════════════════════════

df = df.rename(columns={
    "VDD":          "VDD_ORIG",
    "dnivdd_final": "DNI_CNET",
    "RH.VENDEDOR":  "VENDEDOR",
    "RH.SUPERVISOR":"SUPERVISOR",
    "RH.ESQUEMA":   "ESQUEMA",
    "RH.ZONAL":     "ZONAL",
})

# Columnas derivadas: DNI_INICIAL_VT y DNI vienen de dnivddcnet (≡ DNI_CNET)
df["DNI_INICIAL_VT"] = df["DNI_CNET"]
df["DNI"]            = df["DNI_CNET"]
# VDD_CNET y SUP_CNET son copias de los campos resueltos por RH
df["VDD_CNET"]       = df["VENDEDOR"]
df["SUP_CNET"]       = (df["SUPERVISOR"]
                        .astype(str)
                        .str.upper()
                        .str.replace(r"\s{2,}", " ", regex=True)
                        .str.strip()
                        .str.translate(str.maketrans("ÁÉÍÓÚ", "AEIOU")))
df["SUP1"]           = ""

# ══════════════════════════════════════════════════════════════
# 7. JOIN LCF → RIESG
# ══════════════════════════════════════════════════════════════

URL_LCF = "https://docs.google.com/spreadsheets/d/e/2PACX-1vR_9W58TW-lTrt6_sb4bSgWKkfdcYGV0KXjxKwps0l3qOLGz3MR_eA27hnbAFmBegu85U9jzG5yqe9v/pub?gid=1443303762&single=true&output=csv"

lcf = pd.read_csv(URL_LCF)
lcf["PETICION"] = pd.to_numeric(lcf["PETICION"], errors="coerce").astype("Int64")

df = df.merge(
    lcf[["PETICION", "[Q]"]].rename(columns={"PETICION": "peticion", "[Q]": "RIESG"}),
    on="peticion", how="left"
)
df["RIESG"] = df["RIESG"].fillna(0).astype(int)

# ══════════════════════════════════════════════════════════════
# 8. LIMPIEZA FINAL
# ══════════════════════════════════════════════════════════════

df = df.drop_duplicates(subset="peticion")
df["DESTINOPAQUETE"] = df["DESTINOPAQUETE"].fillna("-")

# ══════════════════════════════════════════════════════════════
# 8b. PIPELINE df_rt (mismo flujo que df pero sin Fecha_Alta filter
#     ni LCF; filtro solo por fecha_registro = @periodo)
# ══════════════════════════════════════════════════════════════

# Limpieza básica
df_rt = df_rt.drop_duplicates(subset="peticion")
for c in ["Fecha_de_alta", "Fecha_Registro", "Fecha_Venta", "Fecha_Alta"]:
    df_rt[c] = pd.to_datetime(df_rt[c], errors="coerce")
df_rt["DNI_ORIG"] = pd.to_numeric(df_rt["DNI_ORIG"], errors="coerce").astype("Int64")
df_rt["peticion"] = pd.to_numeric(df_rt["peticion"], errors="coerce").astype("Int64")
df_rt["FE"]       = df_rt["FE"].astype(str).str.strip()

# Join doble con VENTORY
df_rt = df_rt.merge(ventory_por_fe,  on="FE",      how="left")
df_rt = df_rt.merge(ventory_por_pet, on="peticion", how="left")
df_rt["dnivdd_final"] = df_rt["v1_dnivdd"].combine_first(df_rt["v2_dnivdd"])
df_rt["CTRLNET"]      = df_rt["dnivdd_final"].notna().map({True: "SI", False: "NO"})
df_rt = df_rt.drop(columns=["v1_dnivdd", "v2_dnivdd"])

# Join con RH
df_rt = df_rt.merge(rh_slim, on="dnivdd_final", how="left")

# Columnas calculadas
df_rt["DNI_VENDEDOR"]    = df_rt["dnivdd_final"].combine_first(df_rt["DNI_ORIG"]).astype("Int64")
df_rt["PilotoPR+MONOBA"] = (df_rt["Scoring"] == "FLEX").astype(int)
df_rt["TV"]              = df_rt["sub_producto"].apply(
                               lambda x: "TV" if x in {"TRIO", "DUO INT+TV", "COMPLETA TV"} else "BA")
df_rt["MATCH_DIRECC"]    = df_rt.apply(
    lambda r: None if pd.isna(r["dnivdd_final"])
              else (1 if r["dnivdd_final"] != r["DNI_ORIG"] else 0), axis=1)

# Renombres
df_rt = df_rt.rename(columns={
    "VDD":           "VDD_ORIG",
    "dnivdd_final":  "DNI_CNET",
    "RH.VENDEDOR":   "VENDEDOR",
    "RH.SUPERVISOR": "SUPERVISOR",
    "RH.ESQUEMA":    "ESQUEMA",
    "RH.ZONAL":      "ZONAL",
})

# Columnas derivadas
df_rt["DNI_INICIAL_VT"] = df_rt["DNI_CNET"]
df_rt["DNI"]            = df_rt["DNI_CNET"]
df_rt["VDD_CNET"]       = df_rt["VENDEDOR"]
df_rt["SUP_CNET"]       = (df_rt["SUPERVISOR"]
                           .astype(str)
                           .str.upper()
                           .str.replace(r"\s{2,}", " ", regex=True)
                           .str.strip()
                           .str.translate(str.maketrans("ÁÉÍÓÚ", "AEIOU")))
df_rt["SUP1"]           = ""

# RIESG = 0 para todos (RT no pasa por LCF)
df_rt["RIESG"] = 0

# Limpieza final
df_rt["DESTINOPAQUETE"] = df_rt["DESTINOPAQUETE"].fillna("-")

# ══════════════════════════════════════════════════════════════
# 9. CONSTRUCCIÓN DE TABLAS Y EXPORTACIÓN MULTI-HOJA
# ══════════════════════════════════════════════════════════════

# ── Parámetros de período ─────────────────────────────────────
ayer           = (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")
ayer_ts        = pd.Timestamp(ayer)
inicio_mes     = pd.Timestamp(PERIODO + "-01")
fin_mes        = inicio_mes + pd.offsets.MonthEnd(0)
dias_trans     = max(len(pd.bdate_range(inicio_mes, ayer_ts)), 1)
dias_totales   = len(pd.bdate_range(inicio_mes, fin_mes))

# ── Columnas M-1 y M-2 en altas_historico ────────────────────
col_m1 = "altas_" + (inicio_mes - pd.DateOffset(months=1)).strftime("%Y%m")
col_m2 = "altas_" + (inicio_mes - pd.DateOffset(months=2)).strftime("%Y%m")

# ── Cuotas por vendedor (opcional: cuotas.xlsx con cols DNI, CUOTA) ─
path_cuotas = Path(__file__).parent / "cuotas.xlsx"
if path_cuotas.exists():
    cuotas_df = pd.read_excel(path_cuotas)[["DNI", "CUOTA"]]
    cuotas_df["DNI"] = pd.to_numeric(cuotas_df["DNI"], errors="coerce").astype("Int64")
else:
    print("AVISO: cuotas.xlsx no encontrado — columna CUOTA será 0")
    cuotas_df = pd.DataFrame({"DNI": pd.array([], dtype="Int64"), "CUOTA": pd.array([], dtype=float)})

# ── Columna ZONA en rh ────────────────────────────────────────
zona_col = "ZONAL" if "ZONAL" in rh.columns else "ZONA"
f_ing_col = next((c for c in ["F_INGRESO", "F_ING"] if c in rh.columns), None)

# ── Feriados por mes (Perú) ───────────────────────────────────
FERIADOS_MES = {4:2, 5:1, 6:1, 7:2, 8:1, 9:0, 10:1, 11:1, 12:3}

# ── Antigüedad ────────────────────────────────────────────────
def calc_antiguedad(f_ing):
    if pd.isna(f_ing):
        return None
    dias = (ayer_ts - pd.Timestamp(f_ing)).days
    if dias <= 15:   return "<15d"
    elif dias <= 30: return ">15d"
    elif dias <= 60: return ">30d"
    elif dias <= 90: return ">60d"
    else:            return ">90d"

# ── ZONAL efectivo en df (RH tiene prioridad, fallback al SQL) ─
df["_ZONAL"] = df["ZONAL"].fillna(df["zonal"])
df["_sem"]   = df["Fecha_de_alta"].dt.isocalendar().week.astype(int)

# ── Normalizar RH completa con tipos correctos ────────────────
rh_norm = rh.copy()
rh_norm["DNI"] = pd.to_numeric(rh_norm["DNI"], errors="coerce").astype("Int64")
if f_ing_col:
    rh_norm["F_INGRESO"] = pd.to_datetime(rh_norm[f_ing_col], dayfirst=True, errors="coerce")
else:
    rh_norm["F_INGRESO"] = pd.NaT
rh_norm.rename(columns={zona_col: "ZONAL"}, inplace=True)

# ── Base de vendedores: ACTIVO + EN CAMPO ─────────────────────
_mask = pd.Series([True] * len(rh_norm), index=rh_norm.index)
if "ESTADO" in rh_norm.columns:
    _mask &= rh_norm["ESTADO"].str.upper().str.strip() == "ACTIVO"
if "feedback_rh" in rh_norm.columns:
    _mask &= rh_norm["feedback_rh"].str.upper().str.strip() == "EN CAMPO"

rh_base = (
    rh_norm[_mask][["ZONAL", "SUPERVISOR", "DNI", "VENDEDOR", "ESQUEMA", "F_INGRESO"]]
    .drop_duplicates(subset="DNI")
    .reset_index(drop=True)
)
rh_base["ANTIG"] = rh_base["F_INGRESO"].apply(calc_antiguedad)

# ════════════════════════════════════════════════════════════
# RT / ALTAS — orden de columnas según especificación
# ════════════════════════════════════════════════════════════
# Lista maestra de columnas en el orden requerido
# (incluye ORDEN como posición 3; se calculará por vendedor)
COLS_ORDER = [
    "DNI_VENDEDOR", "Fecha_de_alta", "ORDEN",
    "FE", "peticion", "PilotoPR+MONOBA", "DNI_INICIAL_VT",
    "Q", "producto", "sub_producto", "TV", "CTRLNET", "segmento", "COM",
    "Geografia", "RIESG", "zonal", "DEPARTAMENTO", "DOC_CLIENT", "CLI_EXTRANJ",
    "Canal", "Entidad", "Cod_Punto_Venta", "Punto_Venta", "RUC_Entidad",
    "Razon_Social_Entidad", "CANAL1", "Scoring", "ESQUEMA", "VENDEDOR", "DNI",
    "SUPERVISOR", "VDD_ORIG", "DNI_ORIG", "Zonal_Vendedor", "Segmento_Cliente",
    "Geografia_Cliente", "Region_Cliente", "Zonal_Cliente",
    "Fecha_Registro", "Fecha_Venta", "Fecha_Alta",
    "Cod_PS_Origen", "PS_Origen", "Renta", "Cod_PS_Descuento_Ori", "PS_Descuento_Ori",
    "Renta_Descuento_Ori", "LineaPS_Origen", "LineaNombre_Origen", "LineaRenta_Origen",
    "SpeedyPS_Origen", "SpeedyNombre_Origen", "SpeedyRenta_Origen",
    "CablePS_Origen", "CableNombre_Origen", "CableRenta_Origen",
    "Desc_SVABa_Origen", "Renta_SVABa_Origen", "Cod_SVATv_Origen",
    "Desc_SVATv_Origen", "Renta_SVATv_Origen", "decosventa_renta_origen",
    "decosalquiler_renta_origen", "decoscomodato_renta_origen",
    "dsctopromocional_renta_origen", "internetmovil_renta_origen",
    "moden_renta_origen", "multidestino_renta_origen", "seguridadtotal_renta_origen",
    "ultrawifi_renta_origen", "arpuprincipal_origen", "arpuxsvas_origen",
    "RentaTotal_Origen", "PSPaquete_Renta_Destino", "Cod_PS_Descuento_Des",
    "PS_Descuento_Des", "Renta_Descuento_Des", "LineaPS_Destino", "LineaNombre_Destino",
    "LineaRenta_Destino", "SpeedyPS_Destino", "SpeedyNombre_Destino",
    "SpeedyRenta_Destino", "CablePS_Destino", "CableNombre_Destino",
    "CableRenta_Destino", "Desc_SVABa_Destino", "Renta_SVABa_Destino",
    "Cod_SVATv_Destino", "Desc_SVATv_Destino", "Renta_SVATv_Destino",
    "decosventa_renta_destino", "decosalquiler_renta_destino",
    "decoscomodato_renta_destino", "dsctopromocional_renta_destino",
    "internetmovil_renta_destino", "moden_renta_destino", "multidestino_renta_destino",
    "seguridadtotal_renta_destino", "ultrawifi_renta_destino",
    "arpuprincipal_destino", "arpuxsvas_destino", "saltoxsvas", "saltototal",
    "RentaTotal_Destino", "Ingresos", "Cluster_Origen", "Cluster_Destino",
    "Componentes_Alta", "GAP_ARPU", "Tipo_GAP", "Source_System", "Flag_Duplicado",
    "Flag_Registro", "Flag_Registro_Unico", "Flag_Venta", "Flag_Alta", "Flag_Web",
    "Producto_Web", "Canal_Web", "Flag_Web_Alta", "RU_Primera_Peticion",
    "RU_Primera_Fecha_Registro", "RU_Primer_Canal", "Tecnologia_TV",
    "tecnologia_ba", "velocidad_ba", "Marca_Back", "Telefono_contacto_1",
    "Telefono_contacto_2", "PETICIONES_PRUEBA", "ID_CLIENTE", "CUENTA",
    "NUMERO_PETICION", "ABONADO", "INSCRIPCION", "CUENTAFACTURACIONCD",
    "SEGMENTOCD", "MACROSEGMENTO", "CICLO", "DIRECCIONCD", "SUBLOCALIDADCD",
    "PROVINCIACD", "MUNICIPALIDADCD", "DOCUMENTO", "RUC", "NUMERO_TELEFONO",
    "FEC_REG_PETICION", "FECHACAMBIOESTADO", "ESTADO_PETICION",
    "DESC_ESTADO_PETICION", "ESTADO_SUBPET_ORIGEN", "ESTADO_SUBPET_DESTINO",
    "DESC_ESTADO_SUBPET_DESTINO", "ESTADO_AGRUP_ORIGEN", "ESTADO_AGRUP_DESTINO",
    "DESC_ESTADO_AGRUP_DESTINO", "MOTIVO_ESTADO", "DESC_MOTIVO_ESTADO",
    "MOTIVO_CANCELACION", "DESC_MOTIVO_CANCELACION", "SUBMOTIVO_CANCELACION",
    "DESC_SUBMOTIVO_CANCELACION", "ID_CANAL_VENTA", "ID_PUNTO_VENTA", "ID_VENDEDOR",
    "USUARIO2", "TIP_OPE_CMR_SUBPET_ORIGEN", "TIP_OPE_CMR_SUBPET_DESTINO",
    "TIPO_USO_ORIGEN", "TIPO_USO_DESTINO", "FLAG_BAJA_DEUDA", "LOCALIDAD", "AREA",
    "ORIGENPAQUETE", "DESTINOPAQUETE", "ORIGENLINEA", "DESTINOLINEA",
    "ORIGENINTERNET", "DESTINOINTERNET", "ORIGENCABLE", "DESTINOCABLE",
    "ORIGENBLOQUESTV", "DESTINOBLOQUESTV", "ORIGENDESCUENTO", "DESTINODESCUENTO",
    "ORIGENBLOQUESTVDS", "DESTINOBLOQUESTVDS", "ORIGENMODEM", "DESTINOMODEM",
    "ORIGENDECOVENTA", "DESTINODECOVENTA", "ORIGENDECOALQUILER", "DESTINODECOALQUILER",
    "ORIGENDECOCOMODATO", "DESTINODECOCOMODATO", "ORIGENUWIFI", "DESTINOUWIFI",
    "ORIGENDESCUENTOTEMPORAL", "DESTINODESCUENTOTEMPORAL", "ORIGENINTERNETMOVIL",
    "DESTINOINTERNETMOVIL", "ORIGENMULTIDESTINO", "DESTINOMULTIDESTINO",
    "ORIGENSEGURIDADTOTAL", "DESTINOSEGURIDADTOTAL", "ORIGENPSADMINISTRATIVA",
    "DESTINOPSADMINISTRATIVA", "ORIGENCLUSTER", "DESTINOCLUSTER",
    "ORIGENTECNOLOGIAINTERNET", "DESTINOTECNOLOGIAINTERNET", "ORIGENTECNOLOGIATV",
    "DESTINOTECNOLOGIATV", "ORIGENDECOSIGLAS", "DESTINODECOSIGLAS",
    "REQUE", "ORIGEN_REQUE", "TIPO_REQUE", "MOTIVO_REQUE", "ESTADO_REQUE",
    "OFICINA_ADMINISTRATIVA", "SITUACION_REQUE", "CONDICION_REQUE",
    "FEC_REGISTRO", "FEC_LIQUIDACION", "CLIENTE_CMS", "CUENTA_CMS", "SERVICIO_CMS",
    "CLASE_SERVICIO", "OFERTA_CMS", "USUARIO_CMS", "TELEFONO_VOIP",
    "TIPO_PAQUETE_MULTI", "FECHAFORMULACION", "SEGMENTO_CUENTA", "CMS_CODREQ",
    "CMS_INDORIGREQ", "CMS_TIPREQ", "cms_codsrv", "CMS_CODCLI", "CMS_FECREG",
    "CMS_FECLIQ", "CMS_FECEST", "CMS_FECASG", "CMS_FECPRG", "CMS_CODEDO",
    "CMS_CODSIT", "CMS_CODMOTV", "CMS_CODUSR", "CMS_CODOFEPROD", "CMS_PLAZOCNTR",
    "CMS_DESTIPREQ", "CMS_CODGRPREQ", "CMS_DESGRPREQ", "CMS_COD_PETICION",
    "CMS_NUMTELEFVOIP", "CMS_TIPPAQMUL", "CMS_DISTRITOCD", "CMS_CODNOD",
    "CMS_NROPLANO", "CMS_CODCLASRV", "CMS_DESCLASRV", "CMS_DESTINOTV",
    "CMS_DESTINOBLOQUE", "CMS_DESTINOINTERNET", "CMS_DESTINOVOZIP",
    "CMS_ORIGENTV", "CMS_ORIGENBLOQUE", "CMS_ORIGENINTERNET", "CMS_ORIGENVOZIP",
    "CMS_DESTINOVELOCIDADINTERNET", "CMS_ORIGENVELOCIDADINTERNET",
    "CMS_DESTINODECOALQUILER", "CMS_DESTINODECOVENTA", "CMS_DESTINODECOCOMODATO",
    "CMS_DESTINODECOOTROS", "CMS_ORIGENDECOALQUILER", "CMS_ORIGENDECOVENTA",
    "CMS_ORIGENDECOCOMODATO",
    "CMS_ORIGENMODEM", "CMS_DESTINOMODEM",
    "CMS_DESTINOMACADDRESSDECOALQUILER", "CMS_DESTINOMACADDRESSDECOVENTA",
    "CMS_DESTINOMACADDRESSDECOCOMODATO", "CMS_DESTINOMACADDRESSDECOOTROS",
    "CMS_ORIGENMACADDRESSDECOALQUILER", "CMS_ORIGENMACADDRESSDECOVENTA",
    "CMS_ORIGENMACADDRESSDECOCOMODATO", "CMS_ORIGENMACADDRESSDECOOTROS",
    "CMS_ORIGENMACADDRESSMODEM", "CMS_DESTINMACADDRESSOMODEM",
    "CMS_CODCNLVTA", "CMS_CODGRPVTA", "CMS_CODVDD", "CMS_CODMOTACT",
    "CMS_CODUSRGEN", "CMS_CODCTR", "CMS_DOCUMENTO", "CMS_TIPODOC", "CMS_RUC",
    "CMS_SEGMENTO", "CMS_TELCONTACTO", "MACROSEGMENTOTOP", "GESTIONDECOSTOTAL",
    "GESTIONBLOQUESTV", "GESTIONULTRAWIFI", "TIPOTRANSACCION", "CANALAGRUPADO",
    "SUP1", "CMS_DESC_TIPO_REQUERIMIENTO", "CMS_DESC_SITUACION_REQUERIMIENTO",
    "CMS_DESC_MOTIVO_GENERACION", "CMS_DESC_PRODUCTO", "Region_Vendedor2",
    "MATCH_CONVER", "TRUEFALSE", "Semana", "VDD_CNET", "DNI_CNET",
    "MATCH_DIRECC", "SUP_CNET",
]

# Verificar duplicados en COLS_ORDER (columna duplicada = error de reparación en Excel)
_dupes = [c for c in set(COLS_ORDER) if COLS_ORDER.count(c) > 1]
if _dupes:
    raise ValueError(f"COLS_ORDER tiene columnas duplicadas: {_dupes}")

# ── ALTAS: ordenar df por DNI_VENDEDOR y calcular ORDEN ──────
_base = df.drop(columns=["_ZONAL", "_sem"], errors="ignore")
_base = _base.sort_values("DNI_VENDEDOR", kind="stable").reset_index(drop=True)
_base["ORDEN"] = _base.groupby("DNI_VENDEDOR").cumcount() + 1
cols_presentes_altas = [c for c in COLS_ORDER if c in _base.columns]
_base = _base[cols_presentes_altas]
altas_df = _base.copy()

# ── ALTAS MES ANTERIOR: pipeline simplificado ─────────────────
# VENTORY no existía en el mes anterior; se usa dnivddcnet
# (equivalente a DNI_INICIAL_VT) como aproximación del DNI vendedor.
df_ant = df_ant.drop_duplicates(subset="peticion")
for _c in ["Fecha_de_alta", "Fecha_Registro", "Fecha_Venta", "Fecha_Alta"]:
    df_ant[_c] = pd.to_datetime(df_ant[_c], errors="coerce")
df_ant["DNI_ORIG"]    = pd.to_numeric(df_ant["DNI_ORIG"],    errors="coerce").astype("Int64")
df_ant["dnivddcnet"]  = pd.to_numeric(df_ant["dnivddcnet"],  errors="coerce").astype("Int64")
df_ant["peticion"]    = pd.to_numeric(df_ant["peticion"],    errors="coerce").astype("Int64")

# DNI_VENDEDOR = dnivddcnet (≡ DNI_INICIAL_VT del mes anterior)
df_ant["DNI_VENDEDOR"]    = df_ant["dnivddcnet"]
df_ant["DNI_INICIAL_VT"]  = df_ant["dnivddcnet"]
df_ant["CTRLNET"]         = df_ant["dnivddcnet"].notna().map({True: "SI", False: "NO"})
df_ant["TV"]              = df_ant["sub_producto"].apply(
                                lambda x: "TV" if x in tv_productos else "BA")
df_ant["PilotoPR+MONOBA"] = (df_ant["Scoring"] == "FLEX").astype(int)

# Join LCF → RIESG
df_ant = df_ant.merge(
    lcf[["PETICION", "[Q]"]].rename(columns={"PETICION": "peticion", "[Q]": "RIESG"}),
    on="peticion", how="left"
)
df_ant["RIESG"] = df_ant["RIESG"].fillna(0).astype(int)
df_ant["DESTINOPAQUETE"] = df_ant["DESTINOPAQUETE"].fillna("-")

# Construir altas_ant_df con el mismo orden de columnas
_base_ant = df_ant.drop(columns=["_ZONAL", "_sem"], errors="ignore")
_base_ant = _base_ant.sort_values("DNI_VENDEDOR", kind="stable").reset_index(drop=True)
_base_ant["ORDEN"] = _base_ant.groupby("DNI_VENDEDOR").cumcount() + 1
cols_presentes_ant = [c for c in COLS_ORDER if c in _base_ant.columns]
_base_ant = _base_ant[cols_presentes_ant]
altas_ant_df = _base_ant.copy()

# ── RT: construido desde df_rt (query propio, más registros) ─
_base_rt = df_rt.sort_values("DNI_VENDEDOR", kind="stable").reset_index(drop=True)
# COLS_ORDER sin ORDEN (RT no lleva esa columna)
cols_presentes_rt = [c for c in COLS_ORDER if c != "ORDEN" and c in _base_rt.columns]
rt_df = _base_rt[cols_presentes_rt]

# ════════════════════════════════════════════════════════════
# VDD2 — altas diarias por vendedor
# ════════════════════════════════════════════════════════════
pivot_diario = (
    df.assign(d=df["Fecha_de_alta"].dt.date)
    .groupby(["DNI_VENDEDOR", "d"]).size()
    .unstack(fill_value=0)
)
pivot_diario.columns  = [pd.Timestamp(c) for c in pivot_diario.columns]
pivot_diario.index    = pivot_diario.index.astype("Int64")

vdd2 = rh_base.merge(
    pivot_diario.reset_index().rename(columns={"DNI_VENDEDOR": "DNI"}),
    on="DNI", how="left"
)
date_cols_v2 = sorted([c for c in vdd2.columns if isinstance(c, pd.Timestamp)])
for c in date_cols_v2:
    vdd2[c] = vdd2[c].fillna(0).astype(int)
vdd2["Total general"] = vdd2[date_cols_v2].sum(axis=1)
vdd2 = vdd2[
    ["ZONAL", "SUPERVISOR", "DNI", "VENDEDOR", "ESQUEMA", "ANTIG", "F_INGRESO"]
    + date_cols_v2 + ["Total general"]
].sort_values(["ZONAL", "SUPERVISOR", "VENDEDOR"]).reset_index(drop=True)

# ════════════════════════════════════════════════════════════
# VDD1 — resumen por vendedor
# Base: rh_base (ACTIVO + EN CAMPO), con F_INGRESO garantizado
# ════════════════════════════════════════════════════════════

# Aggregaciones desde la tabla de altas
agg_altas = (
    df.groupby("DNI_CNET")
    .agg(ALTAS=("Q", "sum"), AXB_FXS=("RIESG", "sum"))
    .reset_index()
    .rename(columns={"DNI_CNET": "DNI"})
    .assign(DNI=lambda d: d["DNI"].astype("Int64"))
)

# Tabla MF (nueva fuente): filtrar por MES_INGRESO del período
# ALTAS.MF = count de DNI VDD (cada fila = 1 instalada)
_mes_num = int(PERIODO.split("-")[1])
_anio_num = int(PERIODO.split("-")[0])
mf_filt = mf[
    (mf["MES_INGRESO"] == _mes_num) &
    (mf["AÑO_INGRESO"] == _anio_num)
].copy()
mf_norm = (
    mf_filt.groupby("DNI VDD", as_index=False)
    .size()
    .rename(columns={"DNI VDD": "DNI", "size": "ALTAS.MF"})
)
mf_norm["DNI"] = pd.to_numeric(mf_norm["DNI"], errors="coerce").astype("Int64")

hist_m = hist[["dnivdd", col_m1, col_m2]].rename(
    columns={"dnivdd": "DNI", col_m1: "ALTAS_M-1", col_m2: "ALTAS_M-2"}
)
hist_m["DNI"] = hist_m["DNI"].astype("Int64")

# F_INGRESO desde RH completa (sin filtro) para garantizar que ningún DNI quede vacío
_rh_fingreso = (
    rh_norm[["DNI", "F_INGRESO"]]
    .drop_duplicates(subset="DNI")
    .copy()
)

# Construir VDD1: partir de rh_base (ya filtrado y con F_INGRESO)
# y enriquecer con lookup de F_INGRESO de rh_norm para cubrir vacíos
vdd1 = (
    rh_base                          # ZONAL, SUPERVISOR, DNI, VENDEDOR, ESQUEMA, F_INGRESO, ANTIG
    .merge(agg_altas,    on="DNI", how="left")
    .merge(cuotas_df,    on="DNI", how="left")
    .merge(hist_m,       on="DNI", how="left")
    .merge(mf_norm,      on="DNI", how="left")
    .merge(_rh_fingreso.rename(columns={"F_INGRESO": "_fi"}), on="DNI", how="left")
)

# Garantizar F_INGRESO: usar el de rh_base; si vacío, tomar de rh_norm completa
vdd1["F_INGRESO"] = vdd1["F_INGRESO"].fillna(vdd1["_fi"])
vdd1.drop(columns=["_fi"], inplace=True)

# Recalcular ANTIG sobre F_INGRESO definitivo
vdd1["ANTIG"] = vdd1["F_INGRESO"].apply(calc_antiguedad)

vdd1["ALTAS"]         = vdd1["ALTAS"].fillna(0).astype(int)
vdd1["AXB/FXS"]       = vdd1["AXB_FXS"].fillna(0).astype(int)
vdd1["ALTAS_NETAS"]   = vdd1["ALTAS"] - vdd1["AXB/FXS"]
vdd1["CLUSTER.ALTAS"] = vdd1["ALTAS"].apply(lambda x: ">=7" if x >= 7 else x)

# REG_TOT = SUMIF(RT[DNI_CNET], DNI, RT[Q]) — calculado desde df_rt
_reg_tot = (
    df_rt.groupby("DNI_CNET")["Q"]
    .sum()
    .reset_index()
    .rename(columns={"DNI_CNET": "DNI", "Q": "REG_TOT"})
    .assign(DNI=lambda d: pd.to_numeric(d["DNI"], errors="coerce").astype("Int64"))
)
vdd1 = vdd1.merge(_reg_tot, on="DNI", how="left")
vdd1["REG_TOT"] = vdd1["REG_TOT"].fillna(0).astype(int)

vdd1["%Conver"]       = 0.0   # placeholder; se reemplaza con fórmula Excel al exportar
vdd1["RATIO_CON"]     = 0.0   # placeholder; se reemplaza con fórmula Excel al exportar
vdd1["CUOTA"]         = vdd1["CUOTA"].fillna(8).astype(int)   # default 8 si no está en cuotas.xlsx

# PROYECT.ALT = (ALTAS / dias_lab_transcurridos) * (dias_lab_totales_mes - feriados_mes)
# dias_lab_transcurridos = días hábiles entre min(Fecha_Alta) y max(Fecha_Alta) del periodo
# dias_lab_totales_mes   = días hábiles entre min(Fecha_Alta) y fin de mes
_fechas_alta = df["Fecha_de_alta"].dropna()
if len(_fechas_alta) > 0:
    _f_min = _fechas_alta.min()
    _f_max = _fechas_alta.max()
    _f_fin_mes = _f_min + pd.offsets.MonthEnd(0)
    _mes = _f_min.month
    _feriados = FERIADOS_MES.get(_mes, 0)
    _dias_lab_trans  = max(len(pd.bdate_range(_f_min, _f_max)), 1)
    _dias_lab_totmes = max(len(pd.bdate_range(_f_min, _f_fin_mes)) - _feriados, 1)
else:
    _dias_lab_trans  = 1
    _dias_lab_totmes = 1

vdd1["PROYECT.ALT"]   = (vdd1["ALTAS"] / _dias_lab_trans * _dias_lab_totmes).round(0).astype(int)
vdd1["PROY_VS_CUOTA"] = np.where(
    vdd1["CUOTA"] > 0,
    (vdd1["PROYECT.ALT"] / vdd1["CUOTA"]).round(4),
    np.nan
)
vdd1["ALTAS_M-1"]     = vdd1["ALTAS_M-1"].fillna(0).astype(int)
vdd1["ALTAS_M-2"]     = vdd1["ALTAS_M-2"].fillna(0).astype(int)
vdd1["ALTAS.MF"]      = vdd1["ALTAS.MF"].fillna(0).astype(int)
vdd1["OBSERVACIONES"] = ""

vdd1_final = vdd1[[
    "ZONAL", "SUPERVISOR", "DNI", "VENDEDOR", "ANTIG", "F_INGRESO", "ESQUEMA",
    "ALTAS", "AXB/FXS", "ALTAS_NETAS", "CLUSTER.ALTAS", "REG_TOT", "%Conver", "RATIO_CON",
    "PROYECT.ALT", "CUOTA", "PROY_VS_CUOTA", "ALTAS_M-1", "ALTAS_M-2", "ALTAS.MF",
    "OBSERVACIONES"
]].sort_values(["ZONAL", "SUPERVISOR", "VENDEDOR"]).reset_index(drop=True)

# ════════════════════════════════════════════════════════════
# VDD3 — distribución de vendedores por cantidad de altas
# ════════════════════════════════════════════════════════════
def dist_vdd(df_in):
    def bucket(n):
        if n == 0:   return "0"
        elif n == 1: return "1"
        elif n == 2: return "2"
        elif n == 3: return "3"
        elif n == 4: return "4"
        elif n <= 6: return "6"
        else:        return ">=7"

    tmp = df_in.copy()
    tmp["bucket"] = tmp["ALTAS"].apply(bucket)
    cnt = (tmp.groupby(["ZONAL", "bucket"]).size()
           .unstack(fill_value=0)
           .reindex(columns=["0","1","2","3","4","6",">=7"], fill_value=0))
    cnt["Total [Q]"] = cnt.sum(axis=1)

    rows = []
    for zonal, row in cnt.iterrows():
        entry = {"Etiquetas de fila": zonal}
        for b in ["0","1","2","3","4","6",">=7"]:
            entry[f"{b} [Q]"] = int(row[b])
            entry[f"{b} [%]"] = round(row[b] / row["Total [Q]"], 4) if row["Total [Q]"] > 0 else 0
        entry["Total [Q]"] = int(row["Total [Q]"])
        rows.append(entry)

    total = {"Etiquetas de fila": "Total general"}
    for b in ["0","1","2","3","4","6",">=7"]:
        q = int(cnt[b].sum()); t = int(cnt["Total [Q]"].sum())
        total[f"{b} [Q]"] = q
        total[f"{b} [%]"] = round(q / t, 4) if t > 0 else 0
    total["Total [Q]"] = int(cnt["Total [Q]"].sum())
    rows.append(total)
    return pd.DataFrame(rows)

vdd3_planilla = dist_vdd(vdd1_final[vdd1_final["ESQUEMA"] == "PLANILLA"])
vdd3_nuevos   = dist_vdd(vdd1_final[vdd1_final["ANTIG"].isin(["<15d", ">15d"])])

# ════════════════════════════════════════════════════════════
# TDS — datos base de cuotas leídos desde cuotas_zonal_sup.xlsx
# ════════════════════════════════════════════════════════════

_path_czsup = Path(__file__).parent / "cuotas_zonal_sup.xlsx"
_cz_zon = pd.read_excel(_path_czsup, sheet_name="ZONAL")
_cz_sup = pd.read_excel(_path_czsup, sheet_name="SUPERVISOR")

# Columna de cuota: la segunda columna numérica (CUOTA_ABRIL, CUOTA_MAYO, etc.)
_cuota_col_zon = [c for c in _cz_zon.columns if c != "ZONAL"][0]
_cuota_col_sup = [c for c in _cz_sup.columns if c not in ("ZONAL", "SUPERVISOR")][0]

# Lista cerrada de zonales y cuotas (orden del archivo)
_tds_zonales   = _cz_zon["ZONAL"].tolist()
_tds_cuota_zon = dict(zip(_cz_zon["ZONAL"], _cz_zon[_cuota_col_zon].fillna(0).astype(int)))

# Lista cerrada de supervisores con ZONAL, CUOTA (total), CUOTA_REG y CUOTA_FLEX
_tds_sup_rows  = (_cz_sup[["ZONAL", "SUPERVISOR", _cuota_col_sup, "CUOTA_REG", "CUOTA_FLEX"]]
                  .rename(columns={_cuota_col_sup: "CUOTA"})
                  .fillna(0)
                  .reset_index(drop=True))

# ════════════════════════════════════════════════════════════
# EXPORTAR — libro Excel multi-hoja
# ════════════════════════════════════════════════════════════
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

_dir_salida = Path(__file__).parent / "Archivos_Avance"
_dir_salida.mkdir(exist_ok=True)
ruta = _dir_salida / f"AVANCE_{ayer}.xlsx"

with pd.ExcelWriter(ruta, engine="openpyxl") as writer:

    # ── Fuente base del libro: Aptos Narrow 11 para todas las hojas ──
    from openpyxl.styles import Border, Side
    from openpyxl.utils import column_index_from_string

    _base_font = Font(name="Aptos Narrow", size=11)
    # Modificar el estilo Normal del workbook (herencia para todas las celdas)
    writer.book._named_styles["Normal"].font = _base_font

    # ── Tema "Office" (Office 2016+) ─────────────────────────────────
    # Inyectamos el XML del tema Office estándar de Excel para que el libro
    # use exactamente el tema "Office" y no el tema por defecto de openpyxl.
    _OFFICE_THEME_XML = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">'
        '<a:themeElements>'
        '<a:clrScheme name="Office">'
        '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
        '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
        '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
        '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
        '<a:accent3><a:srgbClr val="A9D18E"/></a:accent3>'
        '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
        '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
        '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
        '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
        '</a:clrScheme>'
        '<a:fontScheme name="Office">'
        '<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
        '<a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        '<a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/>'
        '<a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        '</a:fontScheme>'
        '<a:fmtScheme name="Office">'
        '<a:fillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>'
        '<a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>'
        '<a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
        '</a:fillStyleLst>'
        '<a:lnStyleLst>'
        '<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        '<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        '<a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>'
        '</a:lnStyleLst>'
        '<a:effectStyleLst>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst/></a:effectStyle>'
        '<a:effectStyle><a:effectLst>'
        '<a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">'
        '<a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw>'
        '</a:effectLst></a:effectStyle>'
        '</a:effectStyleLst>'
        '<a:bgFillStyleLst>'
        '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>'
        '<a:gradFill rotWithShape="1"><a:gsLst>'
        '<a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>'
        '<a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>'
        '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
        '</a:bgFillStyleLst>'
        '</a:fmtScheme>'
        '</a:themeElements>'
        '</a:theme>'
    )
    writer.book.loaded_theme = _OFFICE_THEME_XML.encode("utf-8")

    ws_tds = writer.book.create_sheet("TDS", 0)

    # ── helpers de estilo ──────────────────────────────────────────
    _FNT  = "Aptos Narrow"
    _THIN = Side(style="thin")

    def _f(bold=False, size=11, color=None, italic=False):
        kw = dict(name=_FNT, size=size, bold=bold, italic=italic)
        if color: kw["color"] = color
        return Font(**kw)

    def _fill(rgb):
        return PatternFill("solid", fgColor=rgb)

    def _aln(h="center", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    def _bdr(top=False, bottom=False):
        t = _THIN if top    else None
        b = _THIN if bottom else None
        return Border(top=t, bottom=b)

    def _set(ws, row, col, value=None, font=None, fill=None,
             aln=None, fmt=None, border=None):
        c = ws.cell(row=row, column=col, value=value)
        if font:   c.font      = font
        if fill:   c.fill      = fill
        if aln:    c.alignment = aln
        if fmt:    c.number_format = fmt
        if border: c.border    = border
        return c

    def _force_font_ws(ws):
        """Fuerza Aptos Narrow 11 en cada celda de la hoja, preservando bold/italic/color/underline."""
        for row in ws.iter_rows():
            for cell in row:
                f = cell.font
                # Preservar color solo si es tipo RGB (evitar errores con theme colors)
                try:
                    clr = f.color.rgb if f.color and f.color.type == "rgb" else None
                except Exception:
                    clr = None
                cell.font = Font(
                    name="Aptos Narrow", size=11,
                    bold=f.bold, italic=f.italic,
                    color=clr, underline=f.underline,
                )

    # ── colores de encabezado reutilizados ─────────────────────────
    _FILL_VERDE = _fill("C6EFCE")   # verde claro (theme:7 tint:0.8)
    _FILL_NARAN = _fill("FCE4D6")   # naranja claro (theme:6 tint:0.8)
    _FILL_MORAD = _fill("E2EFDA")   # morado claro (theme:8 tint:0.8) — se usa E2EFDA por proximidad
    _FILL_MORAD = _fill("DDD9F7")   # violeta más preciso
    _FILL_GRIS  = _fill("595959")   # gris oscuro (theme:0 tint:-0.25)
    _FILL_AMARI = _fill("FFFF00")   # amarillo brillante
    _FILL_AMPA  = _fill("FFFFCC")   # amarillo pálido
    _FILL_AZUL  = _fill("00B0F0")   # azul cielo
    _FILL_NARAN2= _fill("E97132")   # naranja saturado (theme:6 tint:0)
    _FILL_VIO   = _fill("8EA9C1")   # violeta medio (theme:5 tint:0.4)
    _FILL_ROJO  = _fill("FF9999")   # rojo/rosa (theme:9 tint:0.4)
    _FILL_HDR2  = _fill("F4B183")   # naranja medio (theme:6 tint:0.6) — tabla 3
    _FILL_DORA  = _fill("FFC000")   # amarillo dorado

    # Obtener zonales y supervisores desde datos ya calculados
    _zonales = _tds_zonales
    _nzon    = len(_zonales)           # filas de datos zonal: 6..(5+nzon)
    _sup_rows = _tds_sup_rows          # df: ZONAL, SUPERVISOR, CUOTA

    # ── Anchos de columna ──────────────────────────────────────────
    _col_widths_tds = {
        "A":6.29, "B":16, "C":10, "D":10, "E":11,
        "F":3, "G":3, "H":3, "I":3, "J":3,
        "K":3, "L":3, "M":7.0, "N":9, "O":6.71,
        "P":9.86, "Q":8.71, "R":10, "S":8, "T":8,
        "U":6, "V":7,
        "Y":15, "Z":41, "AA":8.43, "AB":9, "AC":11,
        "AD":3, "AE":3, "AF":3, "AG":3, "AH":3,
        "AI":3, "AJ":3, "AK":6.29, "AL":9, "AM":5.57,
        "AN":11.29, "AO":7.0, "AP":10, "AQ":8, "AR":8,
        "AS":6, "AT":7,
    }
    for col_letter, width in _col_widths_tds.items():
        ws_tds.column_dimensions[col_letter].width = width

    # ── Alturas de fila ────────────────────────────────────────────
    ws_tds.row_dimensions[4].height  = 15.75
    ws_tds.row_dimensions[5].height  = 45.0
    for r in range(6, 6 + _nzon + 12):
        ws_tds.row_dimensions[r].height = 15.0

    # ════════════════════════════════════════════════════
    # FILA 1 — DATE
    # ════════════════════════════════════════════════════
    _set(ws_tds, 1, 1, "DATE:",
         font=_f(bold=True, italic=True, size=10),
         aln=_aln("center", "center"))
    _set(ws_tds, 1, 2, "=MAX(Tbl_ALTAS[Fecha_Alta])",
         font=_f(size=10), fmt="mm-dd-yy", aln=_aln("center", "center"))

    # ════════════════════════════════════════════════════
    # FILA 4 — etiquetas CUOTA DIA (celdas combinadas)
    # ════════════════════════════════════════════════════
    ws_tds.merge_cells("U4:V4")
    _set(ws_tds, 4, column_index_from_string("U"), "CUOTA DIA",
         font=_f(bold=True, size=12), fill=_fill("FFFFCC"),
         aln=_aln("center", "center"))
    ws_tds.merge_cells("AS4:AT4")
    _set(ws_tds, 4, column_index_from_string("AS"), "CUOTA DIA",
         font=_f(bold=True, size=12), fill=_fill("FFFFCC"),
         aln=_aln("center", "center"))

    # ════════════════════════════════════════════════════
    # FILA 5 — Cabeceras TABLA 1 (cols B–V)
    # ════════════════════════════════════════════════════
    _HDR1 = [
        ("B",  "ZONALES",              None,       False),
        ("C",  "CUOTAS ALTAS",         _FILL_VERDE, True),
        ("D",  "AVANCEMES",            _FILL_NARAN, True),
        ("E",  "%AVANCE",              _FILL_NARAN, True),
        ("F",  "=+G5-1",               _FILL_NARAN, False),
        ("G",  "=+H5-1",               _FILL_NARAN, False),
        ("H",  "=+I5-1",               _FILL_NARAN, False),
        ("I",  "=+J5-1",               _FILL_NARAN, False),
        ("J",  "=+K5-1",               _FILL_NARAN, False),
        ("K",  "=+L5-1",               _FILL_NARAN, False),
        ("L",  "=B1",                  _FILL_NARAN, False),
        ("M",  "AXB / FXS",            _FILL_MORAD, True),
        ("N",  "%PROY.",               _FILL_AMARI, True),
        ("O",  "HC",                   _FILL_GRIS,  True),
        ("P",  "HC contrat. Este mes", _FILL_GRIS,  True),
        ("Q",  "Proy. Alt/HC",         _FILL_AMPA,  True),
        ("R",  "%Conver",              _FILL_VERDE, True),
        ("S",  "%FLEX",                _FILL_AZUL,  True),
        ("T",  "ALTAS PDTE",           _FILL_NARAN2,True),
        ("U",  "RUS",                  _FILL_VIO,   True),
        ("V",  "ALTAS",                _FILL_ROJO,  True),
    ]
    for col_l, txt, fill, bold in _HDR1:
        _set(ws_tds, 5, column_index_from_string(col_l), txt,
             font=_f(bold=True, size=11, color="000000"),
             aln=_aln("center", "center", wrap=True),
             border=_bdr(top=True),
             fmt=('ddd\\ dd' if txt.startswith("=") else None))

    # F5:L5 — orientación "Girar el texto hacia arriba" (text_rotation=90)
    for col_l in list("FGHIJKL"):
        ws_tds.cell(row=5, column=column_index_from_string(col_l)).alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True, text_rotation=90
        )

    # ════════════════════════════════════════════════════
    # FILAS 6..(5+nzon) — Datos ZONALES
    # ════════════════════════════════════════════════════
    _r0 = 6   # primera fila de datos
    _r_tot = _r0 + _nzon   # fila TOTAL (AUREN)
    _r_param = _r_tot + 5  # fila 18 de parámetros días lab

    for i, zon in enumerate(_zonales):
        r = _r0 + i
        cuota_zon = int(_tds_cuota_zon.get(zon, 0))
        bdr = _bdr(top=True) if i == 0 else None

        _set(ws_tds, r, column_index_from_string("B"), zon,
             font=_f(bold=True, size=11), aln=_aln("center","center"), border=bdr)
        _set(ws_tds, r, column_index_from_string("C"), cuota_zon,
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0", border=bdr)

        Bc = f"B{r}"; Cc = f"C{r}"; Dc = f"D{r}"
        _set(ws_tds, r, column_index_from_string("D"),
             f'=SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[RIESG],0,Tbl_ALTAS[zonal],TDS!${Bc})',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0", border=bdr)
        _set(ws_tds, r, column_index_from_string("E"),
             f'=+TDS!${Dc}/TDS!${Cc}',
             font=_f(bold=True, size=11), aln=_aln("center","center"), fmt="0%", border=bdr)
        # Días F..L
        for ci, col_l in enumerate(list("FGHIJKL")):
            col_idx = column_index_from_string(col_l)
            head_ref = f"{col_l}$5"
            _set(ws_tds, r, col_idx,
                 f'=SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[RIESG],0,Tbl_ALTAS[Fecha_Alta],TDS!{head_ref},Tbl_ALTAS[zonal],TDS!${Bc})',
                 font=_f(size=11), aln=_aln("center","center"), border=bdr)
        # M AXB/FXS
        _set(ws_tds, r, column_index_from_string("M"),
             f'=SUMIFS(Tbl_ALTAS[RIESG],Tbl_ALTAS[RIESG],1,Tbl_ALTAS[zonal],TDS!${Bc})',
             font=_f(bold=True, size=11), aln=_aln("center","center"), fmt="#,##0", border=bdr)
        # N %PROY
        _set(ws_tds, r, column_index_from_string("N"),
             f'=(TDS!${Dc}/$B${_r_param}*$D${_r_param})/TDS!${Cc}',
             font=_f(bold=True, size=11), aln=_aln("center","center"), fmt="0%", border=bdr)
        # O HC
        _set(ws_tds, r, column_index_from_string("O"),
             f"=SUMIF(Y$2:Y$10000,B{r},AM$2:AM$10000)",
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0", border=bdr)
        # P HC contrat
        _set(ws_tds, r, column_index_from_string("P"),
             f"=SUMIF(Y$2:Y$10000,B{r},AN$2:AN$10000)",
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0", border=bdr)
        # Q Proy Alt/HC
        Oc = f"O{r}"
        _set(ws_tds, r, column_index_from_string("Q"),
             f'=(TDS!${Dc}/$B${_r_param}*$D${_r_param})/{Oc}',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0.0", border=bdr)
        # R %Conver
        _set(ws_tds, r, column_index_from_string("R"),
             f'=IFERROR({Dc}/SUMIF(Tbl_RT[zonal],TDS!{Bc},Tbl_RT[Q]),"")',
             font=_f(bold=True, size=11), aln=_aln("center","center"), fmt="0%", border=bdr)
        # S %FLEX
        _set(ws_tds, r, column_index_from_string("S"),
             f'=IFERROR(SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[zonal],TDS!${Bc},Tbl_ALTAS[Scoring],"FLEX")/{Dc},"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0%", border=bdr)
        # T Altas pendientes
        _set(ws_tds, r, column_index_from_string("T"),
             f'=IF((TDS!${Cc}-TDS!${Dc})<0,0,(TDS!${Cc}-TDS!${Dc}))',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0", border=bdr)
        # V Altas pdte diarias
        Tc = f"T{r}"; Vc = f"V{r}"
        _set(ws_tds, r, column_index_from_string("V"),
             f'=ROUND(TDS!${Tc}/$C${_r_param},0)',
             font=_f(size=11), aln=_aln("center","center"), fmt="0", border=bdr)
        # U RUS cuota día
        _set(ws_tds, r, column_index_from_string("U"),
             f'=IFERROR(IF(ROUND({Vc}/R{r},0)={Vc},{Vc}+1,ROUND({Vc}/R{r},0)),"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0", border=bdr)

    # ════════════════════════════════════════════════════
    # FILA _r_tot — TOTAL AUREN
    # ════════════════════════════════════════════════════
    r  = _r_tot
    r0 = _r0; r1 = _r_tot - 1
    _set(ws_tds, r, column_index_from_string("B"), "AUREN",
         font=_f(bold=True, size=11), aln=_aln("center","center"),
         border=_bdr(top=True, bottom=True))
    _set(ws_tds, r, column_index_from_string("C"), f"=SUBTOTAL(109,C{r0}:C{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0",
         border=_bdr(top=True, bottom=True))
    _set(ws_tds, r, column_index_from_string("D"), f"=SUBTOTAL(109,D{r0}:D{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0",
         border=_bdr(top=True, bottom=True))
    _set(ws_tds, r, column_index_from_string("E"), f"=+TDS!$D{r}/TDS!$C{r}",
         font=_f(bold=True), aln=_aln("center","center"), fmt="0%",
         border=_bdr(top=True, bottom=True))
    for col_l in list("FGHIJKL"):
        _set(ws_tds, r, column_index_from_string(col_l),
             f"=SUM({col_l}{r0}:{col_l}{r1})",
             font=_f(bold=True), aln=_aln("center","center"),
             border=_bdr(top=True, bottom=True))
    _set(ws_tds, r, column_index_from_string("M"), f"=SUBTOTAL(109,M{r0}:M{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0",
         border=_bdr(top=True, bottom=True))
    _set(ws_tds, r, column_index_from_string("N"),
         f"=(TDS!$D{r}/$B${_r_param}*$D${_r_param})/TDS!$C{r}",
         font=_f(bold=True), aln=_aln("center","center"), fmt="0%",
         border=_bdr(top=True, bottom=True))
    _set(ws_tds, r, column_index_from_string("O"), f"=SUM(O{r0}:O{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0",
         border=_bdr(top=True, bottom=True))
    # P13 — mismo formato y fórmula que O13 (HC contratados este mes, total)
    _set(ws_tds, r, column_index_from_string("P"), f"=SUM(P{r0}:P{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0",
         border=_bdr(top=True, bottom=True))
    # Q13 — (D13/$B$18*$D$18)/O13  — Proy Alt/HC total, decimal 1 dec
    _set(ws_tds, r, column_index_from_string("Q"),
         f"=(TDS!$D{r}/$B${_r_param}*$D${_r_param})/O{r}",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0.0",
         border=_bdr(top=True, bottom=True))
    # R13 — D13/SUMA(Tbl_RT[Q])  — %Conver total, porcentaje sin decimales
    _set(ws_tds, r, column_index_from_string("R"),
         f"=IFERROR(D{r}/SUM(Tbl_RT[Q]),\"\")",
         font=_f(bold=True), aln=_aln("center","center"), fmt="0%",
         border=_bdr(top=True, bottom=True))
    # S13 — SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[Scoring],"FLEX",Tbl_ALTAS[RIESG],0)/D13
    _set(ws_tds, r, column_index_from_string("S"),
         f"=IFERROR(SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[Scoring],\"FLEX\",Tbl_ALTAS[RIESG],0)/D{r},\"\")",
         font=_f(bold=True), aln=_aln("center","center"), fmt="0%",
         border=_bdr(top=True, bottom=True))
    # T13 — SUMA(T6:T12)
    _set(ws_tds, r, column_index_from_string("T"), f"=SUM(T{r0}:T{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0",
         border=_bdr(top=True, bottom=True))
    # U13 — SUMA(U6:U12)
    _set(ws_tds, r, column_index_from_string("U"), f"=SUM(U{r0}:U{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="0",
         border=_bdr(top=True, bottom=True))
    # V13 — SUMA(V6:V12)
    _set(ws_tds, r, column_index_from_string("V"), f"=SUM(V{r0}:V{r1})",
         font=_f(bold=True), aln=_aln("center","center"), fmt="0",
         border=_bdr(top=True, bottom=True))

    # B13:V13 — asegurar negrita + borde superior e inferior en todas las columnas
    _bdr_both = _bdr(top=True, bottom=True)
    for _cl in list("BCDEFGHIJKLMNOPQRSTUV"):
        _c = ws_tds.cell(row=r, column=column_index_from_string(_cl))
        _c.font = Font(name=_FNT, size=11, bold=True,
                       color=(_c.font.color if _c.font else None))
        _c.border = _bdr_both

    # ════════════════════════════════════════════════════
    # FILA _r_tot+2 — nota "Basado en..."
    # ════════════════════════════════════════════════════
    _set(ws_tds, _r_tot + 2, column_index_from_string("B"),
         f'="Basado en "&$D${_r_param}&"d lab. y actualizado al "&TEXT($B$1,"dd-mmm")',
         font=_f(bold=True, size=11))

    # ════════════════════════════════════════════════════
    # FILA _r_param — parámetros días laborables
    # ════════════════════════════════════════════════════
    _set(ws_tds, _r_param, column_index_from_string("B"),
         "=NETWORKDAYS.INTL(MIN(Tbl_ALTAS[Fecha_Alta]),MAX(Tbl_ALTAS[Fecha_Alta]),11)",
         font=_f(size=10), aln=_aln("center","center"))
    _set(ws_tds, _r_param, column_index_from_string("C"),
         f"=+D{_r_param}-B{_r_param}",
         font=_f(size=10), aln=_aln("center","center"))
    _set(ws_tds, _r_param, column_index_from_string("D"),
         f"=NETWORKDAYS.INTL(MIN(Tbl_ALTAS[Fecha_Alta]),EOMONTH(MIN(Tbl_ALTAS[Fecha_Alta]),0),11)-{FERIADOS_MES.get(inicio_mes.month, 0)}",
         font=_f(size=10), aln=_aln("center","center"))

    # ════════════════════════════════════════════════════
    # TABLA 2 — Cabeceras SUPERVISOR (cols Y–AT, fila 5)
    # ════════════════════════════════════════════════════
    _HDR2 = [
        ("Y",  "ZONALES",              None,       False),
        ("Z",  "SUPERVISOR",           None,       False),
        ("AA", "CUOTAS ALTAS",         _FILL_VERDE, True),
        ("AB", "AVANCEMES",            _FILL_NARAN, True),
        ("AC", "%AVANCE",              _FILL_NARAN, True),
        ("AD", "=+AE5-1",              _FILL_NARAN, False),
        ("AE", "=+AF5-1",              _FILL_NARAN, False),
        ("AF", "=+AG5-1",              _FILL_NARAN, False),
        ("AG", "=+AH5-1",              _FILL_NARAN, False),
        ("AH", "=+AI5-1",              _FILL_NARAN, False),
        ("AI", "=+AJ5-1",              _FILL_NARAN, False),
        ("AJ", "=B1",                  _FILL_NARAN, False),
        ("AK", "AXB / FXS",            _FILL_MORAD, True),
        ("AL", "%PROY.",               _FILL_AMARI, True),
        ("AM", "HC",                   _FILL_GRIS,  True),
        ("AN", "HC contrat. Este mes", _FILL_GRIS,  True),
        ("AO", "Proy. Alt/HC",         _FILL_AMPA,  True),
        ("AP", "%Conver",              _FILL_VERDE, True),
        ("AQ", "%FLEX",                _FILL_AZUL,  True),
        ("AR", "ALTAS PDTE",           _FILL_NARAN2,True),
        ("AS", "RUS",                  _FILL_VIO,   True),
        ("AT", "ALTAS",                _FILL_ROJO,  True),
    ]
    for col_l, txt, fill, bold in _HDR2:
        _set(ws_tds, 5, column_index_from_string(col_l), txt,
             font=_f(bold=True, size=11, color="000000"),
             aln=_aln("center", "center", wrap=True),
             border=_bdr(top=True),
             fmt=('ddd\\ dd' if txt.startswith("=+") or txt == "=B1" else None))

    # ════════════════════════════════════════════════════
    # FILAS 6..(5+nsup) — Datos SUPERVISORES (Tabla 2)
    # ════════════════════════════════════════════════════
    for i, sup_row in _sup_rows.iterrows():
        r   = _r0 + int(i)
        zon = sup_row["ZONAL"]
        sup = sup_row["SUPERVISOR"]
        cuota_s = int(sup_row["CUOTA"])
        Yc = f"Y{r}"; Zc = f"Z{r}"; AAc = f"AA{r}"; ABc = f"AB{r}"

        _set(ws_tds, r, column_index_from_string("Y"), zon,
             font=_f(size=11), aln=_aln("center","center"))
        _set(ws_tds, r, column_index_from_string("Z"), sup,
             font=_f(size=11), aln=_aln("left","center"))
        _set(ws_tds, r, column_index_from_string("AA"), cuota_s,
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("AB"),
             f'=SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[RIESG],0,Tbl_ALTAS[SUP_CNET],TDS!${Zc})',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("AC"),
             f'=+TDS!${ABc}/TDS!${AAc}',
             font=_f(bold=True), aln=_aln("center","center"), fmt="0%")
        for col_l in ["AD","AE","AF","AG","AH","AI","AJ"]:
            head_ref = f"{col_l}$5"
            _set(ws_tds, r, column_index_from_string(col_l),
                 f'=SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[RIESG],0,Tbl_ALTAS[Fecha_Alta],TDS!{head_ref},Tbl_ALTAS[SUP_CNET],TDS!${Zc})',
                 font=_f(size=11), aln=_aln("center","center"))
        _set(ws_tds, r, column_index_from_string("AK"),
             f'=SUMIFS(Tbl_ALTAS[RIESG],Tbl_ALTAS[RIESG],1,Tbl_ALTAS[SUP_CNET],TDS!${Zc})',
             font=_f(bold=True), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("AL"),
             f'=(TDS!${ABc}/$B${_r_param}*$D${_r_param})/TDS!${AAc}',
             font=_f(bold=True), aln=_aln("center","center"), fmt="0%")
        _set(ws_tds, r, column_index_from_string("AM"),
             f"=COUNTIF('VDD1'!B:B,TDS!{Zc})",
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("AN"),
             f"=COUNTIFS('VDD1'!B:B,TDS!{Zc},'VDD1'!F:F,\">=\"&DATE({_anio_num},{_mes_num},1))",
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        AMc = f"AM{r}"
        _set(ws_tds, r, column_index_from_string("AO"),
             f'=IFERROR((TDS!${ABc}/$B${_r_param}*$D${_r_param})/{AMc},"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0.0")
        _set(ws_tds, r, column_index_from_string("AP"),
             f'=IFERROR(IF(({ABc}/SUMIFS(Tbl_RT[Q],Tbl_RT[SUP_CNET],TDS!${Zc}))>1,1,({ABc}/SUMIFS(Tbl_RT[Q],Tbl_RT[SUP_CNET],TDS!${Zc}))),"")',
             font=_f(bold=True), aln=_aln("center","center"), fmt="0%")
        _set(ws_tds, r, column_index_from_string("AQ"),
             f'=IFERROR(SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[SUP_CNET],TDS!${Zc},Tbl_ALTAS[Scoring],"FLEX")/{ABc},"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0%")
        ARc = f"AR{r}"; ATc = f"AT{r}"
        _set(ws_tds, r, column_index_from_string("AR"),
             f'=IF((TDS!${AAc}-TDS!${ABc})<0,0,(TDS!${AAc}-TDS!${ABc}))',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("AT"),
             f'=ROUND(TDS!${ARc}/$C${_r_param},0)',
             font=_f(size=11), aln=_aln("center","center"), fmt="0")
        _set(ws_tds, r, column_index_from_string("AS"),
             f'=IFERROR(IF(ROUND({ATc}/AP{r},0)={ATc},{ATc}+1,ROUND({ATc}/AP{r},0)),"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0")

    # ════════════════════════════════════════════════════
    # TABLA 3 — por SUPERVISOR (LCF), ahora en AW5
    # ════════════════════════════════════════════════════
    _nsup  = len(_sup_rows)
    _r3_hdr = 5   # fila cabecera fija en fila 5
    ws_tds.row_dimensions[_r3_hdr].height = 45.0

    _HDR3 = [
        ("AW", "ZONALES",          None,       False),
        ("AX", "SUPERVISOR",       None,       False),
        ("AY", "Cuota.REG",        _FILL_HDR2, True),
        ("AZ", "ALTAS",            _FILL_HDR2, True),
        ("BA", "Cuota.FLEX",       _FILL_HDR2, True),
        ("BB", "FLEX",             _FILL_HDR2, True),
        ("BC", "LCF REG",          _FILL_HDR2, True),
        ("BD", "LCF FLEX",         _FILL_HDR2, True),
        ("BE", "ALTAS REG - LCF",  _FILL_HDR2, True),
        ("BF", "ALTAS FLEX - LCF", _FILL_HDR2, True),
        ("BG", "%Cob. REG",        _FILL_HDR2, True),
        ("BH", "%Cob. FLEX",       _FILL_HDR2, True),
        ("BI", "%Cob.REG+FLEX",    _FILL_HDR2, True),
        ("BJ", "Cuota.RUS",        _FILL_DORA, True),
        ("BK", "RU",               _FILL_DORA, True),
        ("BL", "%Cob.",            _FILL_DORA, True),
    ]
    for col_l, txt, fill, bold in _HDR3:
        _set(ws_tds, _r3_hdr, column_index_from_string(col_l), txt,
             font=_f(bold=True, size=11, color="000000"),
             fill=fill,
             aln=_aln("center", "center", wrap=True),
             border=_bdr(bottom=True))

    for i, sup_row in _sup_rows.iterrows():
        r        = _r3_hdr + 1 + int(i)
        zon      = sup_row["ZONAL"]
        sup      = sup_row["SUPERVISOR"]
        cuota_reg  = int(sup_row["CUOTA_REG"])
        cuota_flex = int(sup_row["CUOTA_FLEX"])
        AWc = f"AW{r}"; AXc = f"AX{r}"
        AYc = f"AY{r}"; AZc = f"AZ{r}"; BAc = f"BA{r}"; BBc = f"BB{r}"
        BCc = f"BC{r}"; BDc = f"BD{r}"; BEc = f"BE{r}"; BFc = f"BF{r}"
        BGc = f"BG{r}"; BHc = f"BH{r}"; BJc = f"BJ{r}"; BKc = f"BK{r}"

        _set(ws_tds, r, column_index_from_string("AW"), zon,
             font=_f(size=11), aln=_aln("center","center"))
        _set(ws_tds, r, column_index_from_string("AX"), sup,
             font=_f(size=11), aln=_aln("left","center"))
        # Cuota.REG — desde cuotas_zonal_sup.xlsx campo CUOTA_REG
        _set(ws_tds, r, column_index_from_string("AY"), cuota_reg,
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("AZ"),
             f'=SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[SUP_CNET],TDS!${AXc},Tbl_ALTAS[Scoring],"REGULAR")',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        # Cuota.FLEX — desde cuotas_zonal_sup.xlsx campo CUOTA_FLEX
        _set(ws_tds, r, column_index_from_string("BA"), cuota_flex,
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BB"),
             f'=SUMIFS(Tbl_ALTAS[Q],Tbl_ALTAS[SUP_CNET],TDS!${AXc},Tbl_ALTAS[Scoring],"FLEX")',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BC"),
             f'=SUMIFS(Tbl_ALTAS[RIESG],Tbl_ALTAS[SUP_CNET],TDS!${AXc},Tbl_ALTAS[Scoring],"REGULAR")',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BD"),
             f'=SUMIFS(Tbl_ALTAS[RIESG],Tbl_ALTAS[SUP_CNET],TDS!${AXc},Tbl_ALTAS[Scoring],"FLEX")',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BE"),
             f'={AZc}-{BCc}',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BF"),
             f'={BBc}-{BDc}',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BG"),
             f'=IFERROR({BEc}/{AYc},"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0%")
        _set(ws_tds, r, column_index_from_string("BH"),
             f'=IFERROR({BFc}/{BAc},"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0%")
        _set(ws_tds, r, column_index_from_string("BI"),
             f'=IFERROR(({BEc}+{BFc})/({AYc}+{BAc}),"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0%")
        # Cuota.RUS = REDONDEAR((Cuota.REG + Cuota.FLEX) / 0.8, 0)
        _set(ws_tds, r, column_index_from_string("BJ"),
             f'=ROUND(({AYc}+{BAc})/0.8,0)',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BK"),
             f'=SUMIFS(Tbl_RT[Q],Tbl_RT[SUP_CNET],TDS!${AXc})',
             font=_f(size=11), aln=_aln("center","center"), fmt="#,##0")
        _set(ws_tds, r, column_index_from_string("BL"),
             f'=IFERROR({BKc}/{BJc},"-")',
             font=_f(size=11), aln=_aln("center","center"), fmt="0%")

        # Borde inferior en la última fila
        if int(i) == _nsup - 1:
            for col_l in ["AW","AX","AY","AZ","BA","BB","BC","BD","BE","BF","BG","BH","BI","BJ","BK","BL"]:
                ws_tds.cell(row=r, column=column_index_from_string(col_l)).border = _bdr(bottom=True)

    # ════════════════════════════════════════════════════
    # FORMATOS ADICIONALES TDS
    # ════════════════════════════════════════════════════

    # ── Ocultar líneas de cuadrícula ─────────────────────
    ws_tds.sheet_view.showGridLines = False

    # ── Colores exactos fila 5 (tabla 1: B–V, tabla 2: Y–AT) ─────────
    _F5_COLORS = {
        "C":  "CAEDFB", "D":  "C1F0C8", "E":  "C1F0C8",
        "F":  "C1F0C8", "G":  "C1F0C8", "H":  "C1F0C8",
        "I":  "C1F0C8", "J":  "C1F0C8", "K":  "C1F0C8",
        "L":  "C1F0C8", "M":  "F2CEEF", "N":  "FFFF00",
        "O":  "BFBFBF", "P":  "BFBFBF", "Q":  "FFFF99",
        "R":  "CAEDFB", "S":  "00B0F0", "T":  "CCFF66",
        "U":  "FFC000", "V":  "8ED973",
        "Y":  "FFFFFF", "Z":  "FFFFFF",
        "AA": "CAEDFB",
        "AB": "C1F0C8", "AC": "C1F0C8", "AD": "C1F0C8",
        "AE": "C1F0C8", "AF": "C1F0C8", "AG": "C1F0C8",
        "AH": "C1F0C8", "AI": "C1F0C8", "AJ": "C1F0C8",
        "AK": "F2CEEF", "AL": "FFFF00",
        "AM": "BFBFBF", "AN": "BFBFBF", "AO": "FFFF99",
        "AP": "CAEDFB", "AQ": "00B0F0", "AR": "CCFF66",
        "AS": "FFC000", "AT": "8ED973",
    }
    for _col_l, _rgb in _F5_COLORS.items():
        _c5 = ws_tds.cell(row=5, column=column_index_from_string(_col_l))
        _c5.fill = PatternFill("solid", fgColor=_rgb)
        # Fila 5 completa: negrita, fuente negra
        _c5.font = Font(name=_FNT, size=11, bold=True, color="000000")

    # B5 también negrita negra (sin color de fondo especial)
    _cb5 = ws_tds.cell(row=5, column=column_index_from_string("B"))
    _cb5.font = Font(name=_FNT, size=11, bold=True, color="000000")

    # AD5:AJ5 — orientación "Girar el texto hacia arriba"
    for _col_l in ["AD","AE","AF","AG","AH","AI","AJ"]:
        _c = ws_tds.cell(row=5, column=column_index_from_string(_col_l))
        _c.alignment = Alignment(horizontal="center", vertical="center",
                                 wrap_text=True, text_rotation=90)

    # ── B15 — negrita cursiva azul #0000FF ───────────────────────────
    _c_b15 = ws_tds.cell(row=_r_tot + 2, column=column_index_from_string("B"))
    _c_b15.font = Font(name=_FNT, size=11, bold=True, italic=True, color="0000FF")

    # ── Y17 — negrita cursiva azul #0000FF ───────────────────────────
    _r_b15_y17 = _r_tot + 4
    _c_y17 = ws_tds.cell(row=_r_b15_y17, column=column_index_from_string("Y"))
    _c_y17.font = Font(name=_FNT, size=11, bold=True, italic=True, color="0000FF")

    # ── Y5:AT5 — borde inferior negro ───────────────────────────────
    for _cl_idx in range(column_index_from_string("Y"), column_index_from_string("AT") + 1):
        _c5 = ws_tds.cell(row=5, column=_cl_idx)
        # Preservar borde superior si ya existe, agregar inferior
        _existing = _c5.border
        _c5.border = Border(
            top=_existing.top if _existing.top and _existing.top.style else None,
            bottom=_THIN,
            left=_existing.left if _existing.left and _existing.left.style else None,
            right=_existing.right if _existing.right and _existing.right.style else None,
        )

    # ── Y16:AT16 — borde inferior ────────────────────────────────────
    _r_y16 = _r_tot + 3
    for _cl_idx in range(column_index_from_string("Y"), column_index_from_string("AT") + 1):
        _c = ws_tds.cell(row=_r_y16, column=_cl_idx)
        _c.border = Border(bottom=_THIN)

    # ── AW5:BL5 — negrita y borde inferior negro (cabecera Tabla 3) ──
    for _cl_idx in range(column_index_from_string("AW"), column_index_from_string("BL") + 1):
        _c5aw = ws_tds.cell(row=5, column=_cl_idx)
        _c5aw.font = Font(name=_FNT, size=11, bold=True, color="000000")
        _existing_aw = _c5aw.border
        _c5aw.border = Border(
            top=_existing_aw.top if _existing_aw.top and _existing_aw.top.style else None,
            bottom=_THIN,
            left=_existing_aw.left if _existing_aw.left and _existing_aw.left.style else None,
            right=_existing_aw.right if _existing_aw.right and _existing_aw.right.style else None,
        )

    # ── Formatos condicionales ───────────────────────────────────────
    # Usamos CellIsRule (más fiable que FormulaRule para fill-only en Excel)
    _r_zon_last = _r_tot - 1        # última fila datos zonales (excluye AUREN)
    _r_sup_last = _r0 + _nsup - 1   # última fila datos supervisores

    # F6:L{zon_last} y AD6:AJ{sup_last} — rosa si = 0
    _fill_rosa = PatternFill("solid", fgColor="FFC7CE")
    _font_rosa = Font(name=_FNT, size=11, color="9C0006")
    for _rng_cero in [f"F{_r0}:L{_r_zon_last}", f"AD{_r0}:AJ{_r_sup_last}"]:
        ws_tds.conditional_formatting.add(
            _rng_cero,
            CellIsRule(operator="equal", formula=["0"],
                       fill=_fill_rosa, font=_font_rosa)
        )

    # N6:N{r_tot} y AL6:AL{sup_last} — semáforo aplicado como color fijo
    # (sin formato condicional; xlwings lee los valores calculados y pinta celda a celda)
    # Las referencias se guardan para usarlas en el bloque xlwings al final del script.
    _sem_rango_N  = f"N{_r0}:N{_r_tot}"
    _sem_rango_AL = f"AL{_r0}:AL{_r_sup_last}"

    # ════════════════════════════════════════════════════
    # Forzar fuente Aptos Narrow 11 en TDS (celda a celda)
    # Se hace al final, después de escribir todo el contenido
    _force_font_ws(ws_tds)

    # ── VDD1 ────────────────────────────────────────────────
    _vdd1_export = vdd1_final.copy()
    _vdd1_export.to_excel(writer, sheet_name="VDD1", index=False)

    ws_v1 = writer.sheets["VDD1"]
    nrows = len(_vdd1_export)
    ncols = len(_vdd1_export.columns)

    # ── Colores de encabezado (fila 1) ─────────────────────────
    # Columnas: ZONAL(1) SUPERVISOR(2) DNI(3) VENDEDOR(4) ANTIG(5) F_INGRESO(6)
    #           ESQUEMA(7) ALTAS(8) AXB/FXS(9) ALTAS_NETAS(10) CLUSTER.ALTAS(11)
    #           REG_TOT(12) %Conver(13) RATIO_CON(14) PROYECT.ALT(15) CUOTA(16)
    #           PROY_VS_CUOTA(17) ALTAS_M-1(18) ALTAS_M-2(19) ALTAS.MF(20) OBSERVACIONES(21)
    # Referencia: AVANCE_2026-04-16v5.xlsx
    _V1_HDR = {
        # col: (fgColor, font_color)
        1:  ("4472C4", "FFFFFF"),   # ZONAL          — azul
        2:  ("4472C4", "FFFFFF"),   # SUPERVISOR      — azul
        3:  ("4472C4", "FFFFFF"),   # DNI             — azul
        4:  ("4472C4", "FFFFFF"),   # VENDEDOR        — azul
        5:  ("4472C4", "FFFFFF"),   # ANTIG           — azul
        6:  ("4472C4", "FFFFFF"),   # F_INGRESO       — azul
        7:  ("4472C4", "FFFFFF"),   # ESQUEMA         — azul
        8:  ("4472C4", "FFFFFF"),   # ALTAS           — azul
        9:  ("A02B93", "FFFFFF"),   # AXB/FXS         — morado
        10: ("196B24", "FFFFFF"),   # ALTAS_NETAS     — verde oscuro
        11: ("FFC000", "FFFFFF"),   # CLUSTER.ALTAS   — dorado
        12: ("FFFFFF", "000000"),   # REG_TOT          — blanco, fuente negra
        13: ("FFFFFF", "000000"),   # %Conver          — blanco, fuente negra
        14: ("E97132", "FFFFFF"),   # RATIO_CON        — naranja
        15: ("4472C4", "FFFFFF"),   # PROYECT.ALT     — azul
        16: ("0F9ED5", "FFFFFF"),   # CUOTA            — azul claro
        17: ("4472C4", "FFFFFF"),   # PROY_VS_CUOTA   — azul
        18: ("0F9ED5", "FFFFFF"),   # ALTAS_M-1        — azul claro
        19: ("0F9ED5", "FFFFFF"),   # ALTAS_M-2        — azul claro
        20: ("A02B93", "FFFFFF"),   # ALTAS.MF         — morado
        21: ("FFFF00", "000000"),   # OBSERVACIONES    — amarillo, fuente negra
    }
    # Bandas de datos: filas pares/impares con azul pastel (calculado desde theme accent1 #4472C4)
    _V1_ROW_EVEN = "B4C6E7"   # tint ~0.60
    _V1_ROW_ODD  = "D9E2F3"   # tint ~0.80

    _V1_COL_WIDTHS = {
        1: 11.71, 2: 30.71, 3: 9.0,  4: 40.14, 5: 15.71, 6: 14.29,
        7: 13.43, 8: 10.29, 9: 12.14, 10: 16.0, 11: 17.71, 12: 12.43,
        13: 12.43, 14: 14.57, 15: 16.0, 16: 11.0, 17: 18.71, 18: 14.0,
        19: 14.0,  20: 15.14, 21: 19.0,
    }

    # Aplicar encabezados (fila 1)
    for col_idx in range(1, ncols + 1):
        cell = ws_v1.cell(row=1, column=col_idx)
        bg, fg = _V1_HDR.get(col_idx, ("4472C4", "FFFFFF"))
        cell.font      = Font(name="Aptos Narrow", size=11, bold=True, color=fg)
        cell.fill      = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_v1.column_dimensions[get_column_letter(col_idx)].width = _V1_COL_WIDTHS.get(col_idx, 12)

    # Aplicar bandas de color en filas de datos
    for data_row in range(2, nrows + 2):
        row_fill = PatternFill("solid", fgColor=_V1_ROW_EVEN if data_row % 2 == 0 else _V1_ROW_ODD)
        for col_idx in range(1, ncols + 1):
            cell = ws_v1.cell(row=data_row, column=col_idx)
            cell.fill = row_fill
            cell.font = Font(name="Aptos Narrow", size=11)

    # Formato de datos: columnas con número_format especial
    _DATE_FMT = "dd/mm/yyyy"
    _PCT_FMT  = "0%"
    _DEC1_FMT = "0.0"
    _NUM_FMTS = {6: _DATE_FMT, 13: _PCT_FMT, 14: _DEC1_FMT, 17: _PCT_FMT}
    for col_idx, fmt in _NUM_FMTS.items():
        for row in ws_v1.iter_rows(min_row=2, max_row=nrows + 1,
                                   min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is not None:
                    cell.number_format = fmt

    # Fórmulas en col 13 (%Conver) y col 14 (RATIO_CON)
    _col_dni       = get_column_letter(3)   # C = DNI
    _col_alt_netas = get_column_letter(10)  # J = ALTAS_NETAS
    _col_reg_tot   = get_column_letter(12)  # L = REG_TOT
    for data_row in range(2, nrows + 2):
        an  = f"{_col_alt_netas}{data_row}"
        rt  = f"{_col_reg_tot}{data_row}"
        dni = f"{_col_dni}{data_row}"
        ws_v1.cell(row=data_row, column=13).value         = f"=IFERROR(IF(({an}/{rt})>=1,1,({an}/{rt})),0)"
        ws_v1.cell(row=data_row, column=13).number_format = _PCT_FMT
        ws_v1.cell(row=data_row, column=14).value         = f"=IFERROR(SUMIF(CON[DNIVDD],{dni},CON[Q])/TDS!$B$18,0)"
        ws_v1.cell(row=data_row, column=14).number_format = _DEC1_FMT

    # Formato condicional: AXB/FXS (col I) = 0 → rojo/rosa
    ws_v1.conditional_formatting.add(
        f"I2:I{nrows + 1}",
        CellIsRule(operator="equal", formula=["0"],
                   font=Font(color="9C0006"),
                   fill=PatternFill("solid", fgColor="FFC7CE"))
    )

    # ── VDD2: encabezado + tabla ─────────────────────────────
    vdd2.to_excel(writer, sheet_name="VDD2", index=False, startrow=3)
    ws_v2 = writer.sheets["VDD2"]
    ws_v2["A1"] = "feedback_rh"
    ws_v2["B1"] = "EN CAMPO"
    ws_v2["A3"] = "ALTAS"

    # ── Formato VDD2 ──────────────────────────────────────────
    _v2_nrows = len(vdd2)
    _v2_ncols = len(vdd2.columns)
    _v2_col_names = list(vdd2.columns)
    _HDR_ROW_V2 = 4   # startrow=3 → encabezados en fila 4 (1-based), datos desde fila 5

    # Estilo de encabezados: fondo #156082, fuente blanca, negrita, centrado
    _fill_v2_hdr  = PatternFill("solid", fgColor="156082")
    _font_v2_hdr  = Font(name="Aptos Narrow", size=11, bold=True, color="FFFFFF")
    _aln_center   = Alignment(horizontal="center", vertical="center")
    _aln_left     = Alignment(horizontal="left",   vertical="center")
    _thin_blk     = Side(style="thin", color="000000")
    _bdr_all_blk  = Border(
        left=_thin_blk, right=_thin_blk, top=_thin_blk, bottom=_thin_blk
    )

    # Identificar columna F_INGRESO y columnas de fecha (Timestamps)
    _v2_fi_colidx   = None   # índice 1-based de F_INGRESO
    _v2_date_colidxs = []    # índices 1-based de columnas de fecha (pivot días)
    for _ci, _cname in enumerate(_v2_col_names, start=1):
        if _cname == "F_INGRESO":
            _v2_fi_colidx = _ci
        elif isinstance(_cname, pd.Timestamp):
            _v2_date_colidxs.append(_ci)

    # Encabezados: formato + bordes
    for _ci in range(1, _v2_ncols + 1):
        _cell = ws_v2.cell(row=_HDR_ROW_V2, column=_ci)
        _cell.fill   = _fill_v2_hdr
        _cell.font   = _font_v2_hdr
        _cell.border = _bdr_all_blk
        # Formato personalizado "ddd dd" para columnas de fecha en encabezado
        if isinstance(_v2_col_names[_ci - 1], pd.Timestamp):
            _cell.number_format = "ddd dd"
            _cell.alignment = _aln_center
        else:
            _cell.alignment = _aln_center

    # Datos: bordes en todas las celdas + formatos especiales
    for _row_i in range(_HDR_ROW_V2 + 1, _HDR_ROW_V2 + _v2_nrows + 1):
        for _ci in range(1, _v2_ncols + 1):
            _cell = ws_v2.cell(row=_row_i, column=_ci)
            _cell.border = _bdr_all_blk
            # F_INGRESO → dd/mm/yyyy
            if _ci == _v2_fi_colidx:
                _cell.number_format = "dd/mm/yyyy"
                _cell.alignment = _aln_center

    # Autoajuste de ancho para columnas de fecha (pivot días) y Total general
    _v2_total_colidx = _v2_ncols  # "Total general" es la última columna
    for _ci in _v2_date_colidxs + [_v2_total_colidx]:
        _col_ltr = get_column_letter(_ci)
        # Calcular ancho máximo: encabezado "ddd dd" tiene ~6 chars; datos son enteros
        _max_w = max(
            len("ddd dd") + 1,
            max((len(str(ws_v2.cell(row=_r, column=_ci).value or ""))
                 for _r in range(_HDR_ROW_V2 + 1, _HDR_ROW_V2 + _v2_nrows + 1)),
                default=4)
        )
        ws_v2.column_dimensions[_col_ltr].width = _max_w + 1

    # A3 — fondo amarillo, negrita, centrado
    ws_v2["A3"].fill      = PatternFill("solid", fgColor="FFFF00")
    ws_v2["A3"].font      = Font(name="Aptos Narrow", size=11, bold=True)
    ws_v2["A3"].alignment = _aln_center

    # ── Hoja MiFibra ──────────────────────────────────────────────────────
    # Tres tablas tipo pivot (rangos con formato) a partir de mf_filt.
    # Estructura según imagen:
    #   T1 (A1): filas=fechaInscripcionFicha (día-mes), cols=FILIAL,  vals=sum(Q)
    #   T2 (F1): filas=FILIAL,                          cols=PORTA,   vals=sum(Q)
    #   T3 (K1): filas=FILIAL+RENTA MENSUAL,            cols=PORTA,   vals=sum(Q)
    # Separación: 2 columnas en blanco entre tablas.

    ws_mf = writer.book.create_sheet("MiFibra")
    ws_mf.sheet_view.showGridLines = False

    _FNT_MF   = "Aptos Narrow"
    _BLUE_HDR = "156082"   # fondo cabecera azul oscuro
    _WHITE    = "FFFFFF"
    _GRAY_ROW = "D9E1F2"   # fila total / subtotal
    _BLUE_LGT = "DEEAF1"   # fila alternada opcional

    def _mf_fill(rgb):
        return PatternFill("solid", fgColor=rgb)

    def _mf_font(bold=False, color="000000"):
        return Font(name=_FNT_MF, size=11, bold=bold, color=color)

    def _mf_aln(h="center", v="center"):
        return Alignment(horizontal=h, vertical=v, wrap_text=False)

    _THIN_MF = Side(style="thin", color="000000")
    def _mf_bdr():
        return Border(left=_THIN_MF, right=_THIN_MF,
                      top=_THIN_MF, bottom=_THIN_MF)

    def _write_pivot_range(ws, start_row, start_col, df_pivot, title_label):
        """
        Escribe un pivot (DataFrame con totales) en la hoja ws a partir de
        (start_row, start_col). Devuelve la cantidad de filas escritas.
        Encabezado: fondo azul oscuro, fuente blanca.
        Filas de datos: sin fondo (blanco).
        Fila Total: fondo gris claro, negrita.
        """
        cols = list(df_pivot.columns)   # primera col = etiqueta fila
        n_cols = len(cols)
        r = start_row
        c0 = start_col

        # ── Fila de encabezados ──────────────────────────────────────────
        for ci, col_name in enumerate(cols):
            cell = ws.cell(row=r, column=c0 + ci, value=str(col_name))
            cell.fill      = _mf_fill(_BLUE_HDR)
            cell.font      = _mf_font(bold=True, color=_WHITE)
            cell.alignment = _mf_aln("center")
            cell.border    = _mf_bdr()
        r += 1

        # ── Filas de datos + Total ────────────────────────────────────────
        for row_vals in df_pivot.itertuples(index=False):
            is_total = str(row_vals[0]).strip().lower() in ("total", "total general")
            for ci, val in enumerate(row_vals):
                cell = ws.cell(row=r, column=c0 + ci, value=val)
                cell.border    = _mf_bdr()
                cell.alignment = _mf_aln("right" if ci > 0 else "left")
                if is_total:
                    cell.fill = _mf_fill(_GRAY_ROW)
                    cell.font = _mf_font(bold=True)
                else:
                    cell.font = _mf_font()
            r += 1

        return r - start_row   # total filas escritas (encabezado + datos)

    # ── Filtro de mes del período ─────────────────────────────────────────
    _mf_mes_str = f"{_mes_num:02d}"   # "04" para abril
    # Formatear fecha: "DD-Abr", "DD-May", etc.
    _MESES_CORTO = {
        1:"Ene",2:"Feb",3:"Mar",4:"Abr",5:"May",6:"Jun",
        7:"Jul",8:"Ago",9:"Sep",10:"Oct",11:"Nov",12:"Dic"
    }
    _mes_corto = _MESES_CORTO[_mes_num]

    _mf_work = mf_filt.copy()
    # Parsear fechaInscripcionFicha → fecha completa para formato "ddd dd"
    _mf_work["_fecha"] = pd.to_datetime(
        _mf_work["fechaInscripcionFicha"], dayfirst=True, errors="coerce"
    )
    # Formato "ddd dd": lun 07, mar 08, etc. (locale-independiente en español)
    _DIAS_ES = {0:"lun",1:"mar",2:"mié",3:"jue",4:"vie",5:"sáb",6:"dom"}
    _mf_work["_dia_label"] = _mf_work["_fecha"].apply(
        lambda d: f"{_DIAS_ES[d.weekday()]} {d.day:02d}" if pd.notna(d) else ""
    )
    # Guardar fecha para ordenar filas correctamente
    _mf_work["_dia_orden"] = _mf_work["_fecha"].dt.day
    _mf_work["Q"] = 1  # columna auxiliar para pivots de conteo

    # ── TABLA 1: filas=día (ddd dd), cols=FILIAL, vals=count ─────────────
    _t1 = _mf_work.pivot_table(
        index=["_dia_orden", "_dia_label"], columns="FILIAL", values="Q",
        aggfunc="count", fill_value=0
    ).reset_index()
    _t1.columns.name = None
    _t1 = _t1.drop(columns="_dia_orden")
    # Renombrar ANCASH → CHIMBOTE en T1
    _t1 = _t1.rename(columns={"_dia_label": "Etiquetas de fila", "ANCASH": "CHIMBOTE"})
    _num_cols_t1 = [c for c in _t1.columns if c != "Etiquetas de fila"]
    _t1_total = _t1[_num_cols_t1].sum().to_frame().T
    _t1_total.insert(0, "Etiquetas de fila", "Total general")
    _t1 = pd.concat([_t1, _t1_total], ignore_index=True)
    _t1["Total general"] = _t1[_num_cols_t1].sum(axis=1)

    # ── TABLA 2: filas=FILIAL, cols=PORTA, vals=count ────────────────────
    _t2 = _mf_work.pivot_table(
        index="FILIAL", columns="PORTA", values="Q",
        aggfunc="count", fill_value=0
    ).reset_index()
    _t2.columns.name = None
    # Renombrar ANCASH → CHIMBOTE en T2
    _t2 = _t2.rename(columns={"FILIAL": "Etiquetas de fila"})
    _t2["Etiquetas de fila"] = _t2["Etiquetas de fila"].replace("ANCASH", "CHIMBOTE")
    _num_cols_t2 = [c for c in _t2.columns if c != "Etiquetas de fila"]
    _t2_total = _t2[_num_cols_t2].sum().to_frame().T
    _t2_total.insert(0, "Etiquetas de fila", "Total general")
    _t2 = pd.concat([_t2, _t2_total], ignore_index=True)
    _t2["Total general"] = _t2[_num_cols_t2].sum(axis=1)

    # ── TABLA 3: filas=FILIAL (cabecera) + RENTA MENSUAL (detalle), cols=PORTA
    # Renombrar ANCASH → CHIMBOTE en la fuente antes de construir T3
    _mf_work["FILIAL"] = _mf_work["FILIAL"].replace("ANCASH", "CHIMBOTE")

    # Objetivo: ver instaladas por FILIAL y por RENTA; fila de FILIAL = subtotal,
    # filas de RENTA = detalle. Sin fila "Total general" al final.
    _t3_base = _mf_work.pivot_table(
        index=["FILIAL", "RENTA MENSUAL"], columns="PORTA", values="Q",
        aggfunc="count", fill_value=0
    ).reset_index()
    _t3_base.columns.name = None
    _porta_cols_t3 = [c for c in _t3_base.columns if c not in ("FILIAL", "RENTA MENSUAL")]

    _t3_rows = []
    for _filial, _grp in _t3_base.groupby("FILIAL", sort=True):
        # Fila de subtotal del FILIAL (cabecera del grupo)
        _sub = _grp[_porta_cols_t3].sum()
        _sub_row = {"Etiquetas de fila": _filial, "_es_subtotal": True}
        _sub_row.update(_sub.to_dict())
        _t3_rows.append(_sub_row)
        # Filas de detalle por RENTA MENSUAL
        for _, dr in _grp.sort_values("RENTA MENSUAL").iterrows():
            _det_row = {"Etiquetas de fila": dr["RENTA MENSUAL"], "_es_subtotal": False}
            _det_row.update({c: dr[c] for c in _porta_cols_t3})
            _t3_rows.append(_det_row)

    _t3_built = pd.DataFrame(_t3_rows).reset_index(drop=True)
    _t3_built["Total general"] = _t3_built[_porta_cols_t3].sum(axis=1)
    _t3_is_sub = _t3_built.pop("_es_subtotal")
    _t3_built = _t3_built[["Etiquetas de fila"] + _porta_cols_t3 + ["Total general"]]

    # ── Posiciones de las 3 tablas (col inicio) ───────────────────────────
    _LABEL_T1_COL = 1
    _LABEL_T2_COL = len(_t1.columns) + 3        # +2 cols blanco +1
    _LABEL_T3_COL = _LABEL_T2_COL + len(_t2.columns) + 3

    # ── Títulos en fila 1: amarillo + negrita ─────────────────────────────
    _YELLOW = "FFFF00"
    for _lbl_col, _lbl_txt in [
        (_LABEL_T1_COL, "INSTALADAS"),
        (_LABEL_T2_COL, "INSTALADAS por TIPO"),
        (_LABEL_T3_COL, "INSTALADAS por RENTA"),
    ]:
        _lc = ws_mf.cell(row=1, column=_lbl_col, value=_lbl_txt)
        _lc.fill      = _mf_fill(_YELLOW)
        _lc.font      = _mf_font(bold=True)
        _lc.alignment = _mf_aln("left")

    # ── Función de escritura especializada para T3 (subtotales en gris) ───
    def _write_t3(ws, start_row, start_col, df_data, is_subtotal_series):
        cols = list(df_data.columns)
        r = start_row
        # Encabezado
        for ci, col_name in enumerate(cols):
            cell = ws.cell(row=r, column=start_col + ci, value=str(col_name))
            cell.fill      = _mf_fill(_BLUE_HDR)
            cell.font      = _mf_font(bold=True, color=_WHITE)
            cell.alignment = _mf_aln("center")
            cell.border    = _mf_bdr()
        r += 1
        for row_idx, row_vals in enumerate(df_data.itertuples(index=False)):
            is_sub = bool(is_subtotal_series.iloc[row_idx])
            for ci, val in enumerate(row_vals):
                cell = ws.cell(row=r, column=start_col + ci, value=val)
                cell.border    = _mf_bdr()
                cell.alignment = _mf_aln("right" if ci > 0 else "left")
                if is_sub:
                    cell.fill = _mf_fill(_GRAY_ROW)
                    cell.font = _mf_font(bold=True)
                else:
                    cell.font = _mf_font()
            r += 1

    # ── Escribir las 3 tablas en fila 2 ───────────────────────────────────
    _t1_filas_escritas = _write_pivot_range(ws_mf, 2, _LABEL_T1_COL, _t1, "T1")
    _write_pivot_range(ws_mf, 2, _LABEL_T2_COL, _t2, "T2")
    _write_t3(ws_mf, 2, _LABEL_T3_COL, _t3_built, _t3_is_sub)

    # ── Fila PROYECTADO debajo del Total general de T1 ────────────────────
    # La fila Total general es la última fila escrita por _write_pivot_range.
    # _write_pivot_range devuelve total_filas = encabezado(1) + filas_datos.
    # Fila de inicio de T1 = 2 (encabezado) → datos desde fila 3.
    # Fila Total general = 2 + _t1_filas_escritas - 1
    # Fila PROYECTADO    = 2 + _t1_filas_escritas
    _fila_proy = 2 + _t1_filas_escritas

    # Identificar columnas de T1 en la hoja:
    # col A = "Etiquetas de fila", luego las filiales, luego "Total general"
    # _num_cols_t1 contiene los nombres de columnas numéricas (filiales + Total general)
    # Las columnas de filial están en _num_cols_t1[:-1], Total general en _num_cols_t1[-1]
    _t1_all_cols  = ["Etiquetas de fila"] + list(_t1.columns[1:])  # orden real en hoja
    _filiales_t1  = [c for c in _t1_all_cols if c not in ("Etiquetas de fila", "Total general")]

    # Fila donde está el Total general (para referenciar las celdas con fórmulas)
    _fila_totgen  = _fila_proy - 1

    # Escribir etiqueta
    _lbl = ws_mf.cell(row=_fila_proy, column=_LABEL_T1_COL, value="PROYECTADO")
    _lbl.fill      = _mf_fill(_YELLOW)
    _lbl.font      = _mf_font(bold=True)
    _lbl.alignment = _mf_aln("left")
    _lbl.border    = _mf_bdr()

    # Para cada filial: =REDONDEAR((TotGen_filial / TDS!$B$_r_param) * TDS!$D$_r_param, 0)
    _proy_col_refs = {}  # col_letra → referencia celda para sumar en Total
    for _fi_idx, _fi_nombre in enumerate(_filiales_t1):
        _col_hoja = _LABEL_T1_COL + 1 + _fi_idx   # col en la hoja (1-based)
        _col_letra = get_column_letter(_col_hoja)
        # Celda del Total general de esta filial
        _totgen_cell = f"{_col_letra}{_fila_totgen}"
        _formula = (
            f"=ROUND(({_totgen_cell}/TDS!$B${_r_param})*TDS!$D${_r_param},0)"
        )
        _c = ws_mf.cell(row=_fila_proy, column=_col_hoja, value=_formula)
        _c.fill      = _mf_fill(_YELLOW)
        _c.font      = _mf_font(bold=True)
        _c.alignment = _mf_aln("right")
        _c.border    = _mf_bdr()
        _proy_col_refs[_col_letra] = _c

    # Columna Total general de PROYECTADO = suma de las fórmulas anteriores
    _col_totgen_hoja = _LABEL_T1_COL + len(_t1_all_cols) - 1
    _col_totgen_letra = get_column_letter(_col_totgen_hoja)
    _suma_refs = "+".join(
        f"{get_column_letter(_LABEL_T1_COL + 1 + i)}{_fila_proy}"
        for i in range(len(_filiales_t1))
    )
    _ct = ws_mf.cell(row=_fila_proy, column=_col_totgen_hoja,
                     value=f"={_suma_refs}")
    _ct.fill      = _mf_fill(_YELLOW)
    _ct.font      = _mf_font(bold=True)
    _ct.alignment = _mf_aln("right")
    _ct.border    = _mf_bdr()

    # ── Autofit aproximado de columnas A hasta Q ────────────────────────────
    for _col_cells in ws_mf.columns:
        _col_letter = get_column_letter(_col_cells[0].column)
        if _col_letter <= "Q":
            _max_w = 0
            for _cell in _col_cells:
                if _cell.value:
                    _max_w = max(_max_w, len(str(_cell.value)))
            ws_mf.column_dimensions[_col_letter].width = max(_max_w + 2, 8)

    # ══════════════════════════════════════════════════════════════════════
    # HOJA MOVISTAR
    # ══════════════════════════════════════════════════════════════════════
    ws_mov = writer.book.create_sheet("MOVISTAR")
    ws_mov.sheet_view.showGridLines = False

    # ── Helpers de formato ────────────────────────────────────────────────
    _FNT_MOV  = "Aptos Narrow"
    _THIN_MOV = Side(style="thin", color="000000")
    def _mbdr(top=False, bottom=False, left=False, right=False):
        return Border(
            top    = _THIN_MOV if top    else None,
            bottom = _THIN_MOV if bottom else None,
            left   = _THIN_MOV if left   else None,
            right  = _THIN_MOV if right  else None,
        )
    def _mall_bdr():
        return Border(top=_THIN_MOV, bottom=_THIN_MOV,
                      left=_THIN_MOV, right=_THIN_MOV)
    def _mfont(bold=False, color="000000", size=11):
        return Font(name=_FNT_MOV, bold=bold, color=color, size=size)
    def _mfill(rgb):
        return PatternFill("solid", fgColor=rgb)
    def _maln(h="center", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    # Colores originales
    _MOV_YEL   = "FFD166"  # fondo amarillo dorado — zonales fila 1
    _MOV_DKBL  = "0B3866"  # fuente azul oscuro — texto zonales fila 1
    _MOV_CONV  = "002060"  # fuente azul oscuro — %CONV.
    _MOV_HDR2_BG = "156082"; _MOV_HDR2_FG = "FFFFFF"  # fila 2: azul + blanco
    _MOV_B2_BG   = "156082"; _MOV_B2_FG   = "FFFFFF"  # bloques 2: azul + blanco
    _MOV_B3_BG   = "156082"; _MOV_B3_FG   = "FFFFFF"  # bloques 3: azul + blanco
    _MOV_B4_BG   = "156082"; _MOV_B4_FG   = "FFFFFF"  # bloques 4: azul + blanco

    # Zonales en orden alfabético
    _MOV_ZONALES = sorted(_tds_zonales)   # ['AREQUIPA','CHIMBOTE','HUARAZ','ILO','NORTE CHICO','TACNA','TRUJILLO']
    _NZ = len(_MOV_ZONALES)

    # Inicio del período y mes anterior
    _mov_inicio = pd.Timestamp(f"{PERIODO}-01")
    _mov_fin    = _mov_inicio + pd.offsets.MonthEnd(0)
    _dias_mes   = pd.date_range(_mov_inicio, _mov_fin, freq="D")
    _dias_labels = [d.strftime("%a %d").lower().replace("mon","lun").replace("tue","mar")\
                    .replace("wed","mié").replace("thu","jue").replace("fri","vie")\
                    .replace("sat","sáb").replace("sun","dom") for d in _dias_mes]

    _mov_ant_inicio = _mov_inicio - pd.offsets.MonthBegin(1)
    _mov_ant_fin    = _mov_inicio - pd.timedelta_range(start="1 day", periods=1)[0]
    _dias_ant       = pd.date_range(_mov_ant_inicio, _mov_ant_fin, freq="D")

    # ── Preparar datos base ───────────────────────────────────────────────
    # altas_df y rt_df ya están en scope
    _altas  = altas_df.copy()
    _rt     = rt_df.copy()

    # Normalizar fechas a date
    _altas["_fa"]  = pd.to_datetime(_altas["Fecha_Alta"],     errors="coerce").dt.normalize()
    _rt["_fr"]     = pd.to_datetime(_rt["Fecha_Registro"],    errors="coerce").dt.normalize()

    # Filtros mes actual (RIESG==0)
    _altas_mes  = _altas[(_altas["_fa"] >= _mov_inicio) & (_altas["_fa"] <= _mov_fin) & (_altas["RIESG"] == 0)]
    _rt_mes     = _rt[   (_rt["_fr"]   >= _mov_inicio) & (_rt["_fr"]   <= _mov_fin)   & (_rt["RIESG"]   == 0)]

    # Mes anterior: usar altas_ant_df (datos SQL con PERIODO_ANT), RIESG==0
    _altas_ant_raw = altas_ant_df.copy()
    _altas_ant_raw["_fa"] = pd.to_datetime(_altas_ant_raw["Fecha_Alta"], errors="coerce").dt.normalize()
    _altas_ant = _altas_ant_raw[_altas_ant_raw["RIESG"] == 0]

    # ── Función: pivot diario por zonal ───────────────────────────────────
    def _pivot_dia_zon(df, date_col, val_col, agg, dias, zonales, filtro=None):
        """Retorna DataFrame: index=día (date), columns=zonal, values=agg."""
        _d = df.copy()
        if filtro is not None:
            _d = _d[filtro(_d)]
        # Normalizar nombre del campo zonal (puede ser "zonal" o "ZONAL")
        if "ZONAL" in _d.columns and "zonal" not in _d.columns:
            _d = _d.rename(columns={"ZONAL": "zonal"})
        _d["_dia"] = _d[date_col].dt.normalize()
        if agg == "sum":
            _piv = _d.groupby(["_dia", "zonal"])[val_col].sum().unstack(fill_value=0)
        elif agg == "nunique":
            _piv = _d.groupby(["_dia", "zonal"])[val_col].nunique().unstack(fill_value=0)
        _piv = _piv.reindex(index=dias, columns=zonales, fill_value=0)
        return _piv

    # ── Función genérica de escritura de bloque ───────────────────────────
    def _write_mov_block(ws, r_start, c_start, dias, dias_labels, zon_list,
                         data_cols, col_headers, row_total, row_label="Total",
                         hdr_fill="156082", extra_rows=None):
        """
        Escribe un bloque con:
          col 0: Fecha/DIA
          col 1..N: una columna por cada item en col_headers
          col N+1: Total general
        data_cols: lista de pd.Series o arrays alineados con dias (una por col_header)
        extra_rows: lista de (label, [val_por_col, ...]) para filas adicionales debajo de Total
        Retorna fila siguiente al último escrito.
        """
        _nc = len(col_headers)
        c0  = c_start

        # Encabezado
        _hdr_cells = ["Fecha"] + list(col_headers) + ["Total general"]
        for ci, hdr in enumerate(_hdr_cells):
            cell = ws.cell(row=r_start, column=c0+ci, value=hdr)
            cell.fill      = _mfill(hdr_fill)
            cell.font      = _mfont(bold=True, color="FFFFFF")
            cell.alignment = _maln("center")
            cell.border    = _mall_bdr()
        r = r_start + 1

        # Filas de datos
        for di, (dia, lbl) in enumerate(zip(dias, dias_labels)):
            vals = [int(dc.iloc[di]) if hasattr(dc,"iloc") else int(dc[di]) for dc in data_cols]
            row_data = [lbl] + vals + [sum(vals)]
            for ci, v in enumerate(row_data):
                cell = ws.cell(row=r, column=c0+ci, value=v)
                cell.font      = _mfont()
                cell.alignment = _maln("center" if ci > 0 else "left")
                cell.border    = _mall_bdr()
                if ci == 0:
                    cell.number_format = "@"
            r += 1

        # Fila Total
        tot_vals = [int(dc.sum()) if hasattr(dc,"sum") else sum(dc) for dc in data_cols]
        tot_row  = [row_label] + tot_vals + [sum(tot_vals)]
        for ci, v in enumerate(tot_row):
            cell = ws.cell(row=r, column=c0+ci, value=v)
            cell.font      = _mfont(bold=True)
            cell.alignment = _maln("center" if ci > 0 else "left")
            cell.border    = _mall_bdr()
        r_total_row = r
        r += 1

        # Filas extra (Proyectado, Cuota, %Avance…)
        if extra_rows:
            for (ex_label, ex_vals) in extra_rows:
                ex_row = [ex_label] + list(ex_vals) + [sum(v for v in ex_vals if isinstance(v,(int,float)))]
                for ci, v in enumerate(ex_row):
                    cell = ws.cell(row=r, column=c0+ci, value=v)
                    cell.font      = _mfont(bold=(ci==0))
                    cell.alignment = _maln("center" if ci > 0 else "left")
                    cell.border    = _mall_bdr()
                r += 1

        return r, r_total_row

    # ════════════════════════════════════════════════════════════════
    # BLOQUE 1 — CONVERSIÓN (REG_TOT / ALTAS / %CONV. por día/zonal)
    # Fila 1: encabezados de ZONAL (cada grupo ocupa 3 cols: REG/ALT/%)
    # Fila 2: sub-encabezados REG_TOT / ALTAS / %CONV.
    # Filas 3+: datos diarios
    # Fila Total: sumas + %conv total
    # ════════════════════════════════════════════════════════════════

    # Pivots diarios
    _reg_piv   = _pivot_dia_zon(_rt,    "_fr", "Q",           "sum",     _dias_mes, _MOV_ZONALES)
    _alta_piv  = _pivot_dia_zon(_altas_mes, "_fa", "Q",       "sum",     _dias_mes, _MOV_ZONALES)

    _R1 = 1    # fila inicio bloque 1
    _CA = 1    # col A

    # Fila 1: nombres de ZONAL — "Centrar en selección" sobre sus 3 cols (REG/ALT/%)
    # openpyxl: horizontal="centerContinuous" aplica "Centrar en selección"
    _zon_positions = {}   # zonal -> col inicio (1-based)
    for zi, zon in enumerate(_MOV_ZONALES):
        col_z = _CA + 1 + zi * 3
        _zon_positions[zon] = col_z
        # Primera celda del grupo: lleva el valor y el centrado continuo
        cell = ws_mov.cell(row=_R1, column=col_z, value=zon)
        cell.fill      = _mfill(_MOV_YEL)
        cell.font      = _mfont(bold=True, color=_MOV_DKBL)
        cell.alignment = Alignment(horizontal="centerContinuous", vertical="center")
        cell.border    = _mall_bdr()
        # Celdas 2ª y 3ª del grupo: vacías, mismo fondo, mismo centrado continuo
        for _dc in [1, 2]:
            _ec = ws_mov.cell(row=_R1, column=col_z+_dc)
            _ec.fill      = _mfill(_MOV_YEL)
            _ec.alignment = Alignment(horizontal="centerContinuous", vertical="center")
            _ec.border    = _mall_bdr()

    # AUREN (total) — misma lógica
    _col_auren = _CA + 1 + _NZ * 3
    cell = ws_mov.cell(row=_R1, column=_col_auren, value="AUREN")
    cell.fill      = _mfill(_MOV_YEL)
    cell.font      = _mfont(bold=True, color=_MOV_DKBL)
    cell.alignment = Alignment(horizontal="centerContinuous", vertical="center")
    cell.border    = _mall_bdr()
    for _dc in [1, 2]:
        _ec = ws_mov.cell(row=_R1, column=_col_auren+_dc)
        _ec.fill      = _mfill(_MOV_YEL)
        _ec.alignment = Alignment(horizontal="centerContinuous", vertical="center")
        _ec.border    = _mall_bdr()

    # col A fila 1 vacía con fondo amarillo
    _c_a1 = ws_mov.cell(row=_R1, column=_CA)
    _c_a1.fill = _mfill(_MOV_YEL)
    _c_a1.border = _mall_bdr()

    # Fila 2: sub-encabezados
    _c_a2 = ws_mov.cell(row=_R1+1, column=_CA, value="Fecha")
    _c_a2.fill = _mfill(_MOV_YEL); _c_a2.font = _mfont(bold=True, color=_MOV_DKBL)
    _c_a2.alignment = _maln("center"); _c_a2.border = _mall_bdr()

    for zi in range(_NZ + 1):   # +1 para AUREN
        col_z = _CA + 1 + zi * 3
        for ci, (hdr, hfill, hfont) in enumerate([
            ("REG_TOT", _MOV_HDR2_BG, _MOV_HDR2_FG),  # fondo blanco, texto negro
            ("ALTAS",   _MOV_HDR2_BG, _MOV_HDR2_FG),  # fondo blanco, texto negro
            ("%CONV.",  None,          _MOV_CONV),      # sin fondo, texto azul oscuro
        ]):
            cell = ws_mov.cell(row=_R1+1, column=col_z+ci, value=hdr)
            if hfill:
                cell.fill = _mfill(hfill)
            cell.font      = _mfont(bold=True, color=hfont)
            cell.alignment = _maln("center")
            cell.border    = _mall_bdr()

    # Función de fórmula %CONV
    def _pct_conv_formula(alt_col_l, reg_col_l, row):
        return f'=IFERROR({alt_col_l}{row}/{reg_col_l}{row},"")'

    # Filas de datos diarios
    _r_data_start = _R1 + 2
    for di, (dia, lbl) in enumerate(zip(_dias_mes, _dias_labels)):
        r = _r_data_start + di
        # col A: fecha con formato ddd dd
        cell = ws_mov.cell(row=r, column=_CA, value=lbl)
        cell.font = _mfont(); cell.alignment = _maln("center"); cell.border = _mall_bdr()

        for zi, zon in enumerate(_MOV_ZONALES):
            col_z = _CA + 1 + zi * 3
            reg_v  = int(_reg_piv.loc[dia, zon])  if zon in _reg_piv.columns  else 0
            alt_v  = int(_alta_piv.loc[dia, zon]) if zon in _alta_piv.columns else 0
            c_reg = ws_mov.cell(row=r, column=col_z,   value=reg_v)
            c_reg.font = _mfont(); c_reg.alignment = _maln("center"); c_reg.border = _mall_bdr()
            c_alt = ws_mov.cell(row=r, column=col_z+1, value=alt_v)
            c_alt.font = _mfont(); c_alt.alignment = _maln("center"); c_alt.border = _mall_bdr()
            c_pct = ws_mov.cell(row=r, column=col_z+2,
                value=_pct_conv_formula(get_column_letter(col_z+1), get_column_letter(col_z), r))
            c_pct.number_format = "0%"; c_pct.font = _mfont()
            c_pct.alignment = _maln("center"); c_pct.border = _mall_bdr()

        # AUREN (sumas de todos los zonales)
        col_z = _col_auren
        _reg_cols = "+".join(get_column_letter(_CA+1+zi*3) + str(r) for zi in range(_NZ))
        _alt_cols = "+".join(get_column_letter(_CA+2+zi*3) + str(r) for zi in range(_NZ))
        c_ar = ws_mov.cell(row=r, column=col_z,   value=f"={_reg_cols}")
        c_ar.font = _mfont(); c_ar.alignment = _maln("center"); c_ar.border = _mall_bdr()
        c_aa = ws_mov.cell(row=r, column=col_z+1, value=f"={_alt_cols}")
        c_aa.font = _mfont(); c_aa.alignment = _maln("center"); c_aa.border = _mall_bdr()
        c_ap = ws_mov.cell(row=r, column=col_z+2,
            value=_pct_conv_formula(get_column_letter(col_z+1), get_column_letter(col_z), r))
        c_ap.number_format = "0%"; c_ap.font = _mfont(color=_MOV_CONV)
        c_ap.alignment = _maln("center"); c_ap.border = _mall_bdr()

    # Fila Total (fila 11 en el modelo = r_data_start + n_dias)
    _r_tot1 = _r_data_start + len(_dias_mes)
    cell = ws_mov.cell(row=_r_tot1, column=_CA, value="Total")
    cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
    _r0d = _r_data_start; _r1d = _r_tot1 - 1
    for zi in range(_NZ + 1):
        col_z = _CA + 1 + zi * 3
        c_sr = ws_mov.cell(row=_r_tot1, column=col_z,
                           value=f"=SUM({get_column_letter(col_z)}{_r0d}:{get_column_letter(col_z)}{_r1d})")
        c_sr.font = _mfont(bold=True); c_sr.alignment = _maln("center"); c_sr.border = _mall_bdr()
        c_sa = ws_mov.cell(row=_r_tot1, column=col_z+1,
                           value=f"=SUM({get_column_letter(col_z+1)}{_r0d}:{get_column_letter(col_z+1)}{_r1d})")
        c_sa.font = _mfont(bold=True); c_sa.alignment = _maln("center"); c_sa.border = _mall_bdr()
        c_sp = ws_mov.cell(row=_r_tot1, column=col_z+2,
                           value=f"={get_column_letter(col_z+1)}{_r_tot1}/{get_column_letter(col_z)}{_r_tot1}")
        c_sp.number_format = "0%"; c_sp.font = _mfont(bold=True)
        c_sp.alignment = _maln("center"); c_sp.border = _mall_bdr()

    # ════════════════════════════════════════════════════════════════
    # BLOQUE 2 — ALTAS TOTALES / REGULARES / FLEX
    # Fila _R2: título; Fila _R2+1: encabezados; Fila _R2+2+: datos
    # 3 subtablas: A–I (TOTALES), K–S (REGULARES), U–AC (FLEX)
    # Col J y T vacías (separadoras)
    # ════════════════════════════════════════════════════════════════
    _R2 = _r_tot1 + 6   # deja ~5 filas de margen (Proyectado, Cuota, %Avance + blancos)

    # Proyectado (fila _r_tot1 + 1)
    _r_proy1 = _r_tot1 + 1
    cell = ws_mov.cell(row=_r_proy1, column=_CA, value="Proyectado")
    cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
    for zi in range(_NZ + 1):
        col_z = _CA + 1 + zi * 3
        _alt_col_l = get_column_letter(col_z + 1)
        _reg_col_l = get_column_letter(col_z)
        # Proyectado = ROUND((ALTAS_Total / dias_transcurridos) * dias_totales_mes, 0)
        _f = f"=ROUND(({_alt_col_l}{_r_tot1}/TDS!$B${_r_param})*TDS!$D${_r_param},0)"
        c = ws_mov.cell(row=_r_proy1, column=col_z+1, value=_f)
        c.font = _mfont(); c.alignment = _maln("center"); c.border = _mall_bdr()
        # REG vacío, %CONV vacío
        ws_mov.cell(row=_r_proy1, column=col_z).border   = _mall_bdr()
        ws_mov.cell(row=_r_proy1, column=col_z+2).border = _mall_bdr()

    # Cuota (fila _r_tot1 + 2)
    _r_cuota1 = _r_tot1 + 2
    cell = ws_mov.cell(row=_r_cuota1, column=_CA, value="Cuota")
    cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
    for zi, zon in enumerate(_MOV_ZONALES):
        col_z = _CA + 1 + zi * 3
        cuota_v = _tds_cuota_zon.get(zon, 0)
        c = ws_mov.cell(row=_r_cuota1, column=col_z+1, value=cuota_v)
        c.font = _mfont(); c.alignment = _maln("center"); c.border = _mall_bdr()
        ws_mov.cell(row=_r_cuota1, column=col_z).border   = _mall_bdr()
        ws_mov.cell(row=_r_cuota1, column=col_z+2).border = _mall_bdr()
    # AUREN cuota total
    col_z = _col_auren
    _auren_cuota_cols = "+".join(
        get_column_letter(_CA+2+zi*3)+str(_r_cuota1) for zi in range(_NZ)
    )
    c = ws_mov.cell(row=_r_cuota1, column=col_z+1, value=f"={_auren_cuota_cols}")
    c.font = _mfont(); c.alignment = _maln("center"); c.border = _mall_bdr()

    # %Avance (fila _r_tot1 + 3)
    _r_av1 = _r_tot1 + 3
    cell = ws_mov.cell(row=_r_av1, column=_CA, value="%Avance")
    cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
    for zi in range(_NZ + 1):
        col_z = _CA + 1 + zi * 3
        _alt_l = get_column_letter(col_z+1)
        c = ws_mov.cell(row=_r_av1, column=col_z+1,
                        value=f"={_alt_l}{_r_tot1}/{_alt_l}{_r_cuota1}")
        c.number_format = "0%"; c.font = _mfont()
        c.alignment = _maln("center"); c.border = _mall_bdr()
        ws_mov.cell(row=_r_av1, column=col_z).border   = _mall_bdr()
        ws_mov.cell(row=_r_av1, column=col_z+2).border = _mall_bdr()

    # ── Subtablas TOTALES / REGULARES / FLEX ─────────────────────────────
    # Columnas de inicio: A=1, K=11, U=21
    _R2_HDR_TITLE = _R2     # títulos "ALTAS TOTALES" etc
    _R2_HDR_COL   = _R2 + 1
    _R2_DATA      = _R2 + 2

    # Pivots: ALTAS TOTALES (todas RIESG==0), REGULARES (Scoring!='FLEX'), FLEX (Scoring=='FLEX')
    def _alta_pivot_dia(filtro_fn, dias):
        _d = _altas_mes[filtro_fn(_altas_mes)].copy() if filtro_fn else _altas_mes.copy()
        if "ZONAL" in _d.columns and "zonal" not in _d.columns:
            _d = _d.rename(columns={"ZONAL": "zonal"})
        _p = _d.groupby([_d["_fa"].dt.normalize(), "zonal"])["Q"].sum().unstack(fill_value=0)
        return _p.reindex(index=dias, columns=_MOV_ZONALES, fill_value=0)

    _piv_tot  = _alta_pivot_dia(None,                                      _dias_mes)
    _piv_reg  = _alta_pivot_dia(lambda d: d["Scoring"] != "FLEX",          _dias_mes)
    _piv_flex = _alta_pivot_dia(lambda d: d["Scoring"] == "FLEX",          _dias_mes)

    _SUBTABLAS = [
        (1,  "ALTAS TOTALES",   _piv_tot,  None),
        (11, "ALTAS REGULARES", _piv_reg,  "reg"),
        (21, "ALTAS FLEX",      _piv_flex, "flex"),
    ]

    _r_tot2 = {}   # guarda fila Total de cada subtabla para las filas %REGULAR/%FLEX

    for _c_ini, _titulo, _piv, _key in _SUBTABLAS:
        # Título
        cell = ws_mov.cell(row=_R2_HDR_TITLE, column=_c_ini, value=_titulo)
        cell.font = _mfont(bold=True, size=11)
        cell.alignment = _maln("left")

        # Encabezados de columna
        _hdrs2 = ["Fecha"] + _MOV_ZONALES + ["Total general"]
        for ci, hdr in enumerate(_hdrs2):
            cell = ws_mov.cell(row=_R2_HDR_COL, column=_c_ini+ci, value=hdr)
            cell.fill      = _mfill(_MOV_B2_BG)
            cell.font      = _mfont(bold=True, color=_MOV_B2_FG)
            cell.alignment = _maln("center")
            cell.border    = _mall_bdr()

        # Datos diarios
        for di, (dia, lbl) in enumerate(zip(_dias_mes, _dias_labels)):
            r = _R2_DATA + di
            vals = [int(_piv.loc[dia, zon]) if zon in _piv.columns else 0
                    for zon in _MOV_ZONALES]
            row_data = [lbl] + vals + [sum(vals)]
            for ci, v in enumerate(row_data):
                cell = ws_mov.cell(row=r, column=_c_ini+ci, value=v)
                cell.font      = _mfont()
                cell.alignment = _maln("center" if ci > 0 else "left")
                cell.border    = _mall_bdr()

        # Fila Total general
        _r_tg = _R2_DATA + len(_dias_mes)
        _r_tot2[_key] = _r_tg
        cell = ws_mov.cell(row=_r_tg, column=_c_ini, value="Total general")
        cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
        for ci in range(_NZ + 1):
            _cl  = get_column_letter(_c_ini + 1 + ci)
            c_t  = ws_mov.cell(row=_r_tg, column=_c_ini+1+ci,
                               value=f"=SUM({_cl}{_R2_DATA}:{_cl}{_r_tg-1})")
            c_t.font = _mfont(bold=True); c_t.alignment = _maln("center"); c_t.border = _mall_bdr()

    # Fila %REGULAR (debajo del Total de REGULARES, col 11)
    _r_preg = _r_tot2["reg"] + 1
    cell = ws_mov.cell(row=_r_preg, column=11, value="%REGULAR")
    cell.font = _mfont(bold=True); cell.alignment = _maln("left"); cell.border = _mall_bdr()
    for ci in range(_NZ + 1):
        _col_reg = get_column_letter(12 + ci)
        _col_tot = get_column_letter(2  + ci)
        _r_tg_reg = _r_tot2["reg"]; _r_tg_tot = _r_tot2[None]
        c = ws_mov.cell(row=_r_preg, column=12+ci,
                        value=f"=IFERROR({_col_reg}{_r_tg_reg}/{_col_tot}{_r_tg_tot},\"\")")
        c.number_format = "0%"; c.font = _mfont()
        c.alignment = _maln("center"); c.border = _mall_bdr()

    # Fila %FLEX (debajo del Total de FLEX, col 21)
    _r_pflex = _r_tot2["flex"] + 1
    cell = ws_mov.cell(row=_r_pflex, column=21, value="%FLEX")
    cell.font = _mfont(bold=True); cell.alignment = _maln("left"); cell.border = _mall_bdr()
    for ci in range(_NZ + 1):
        _col_flex = get_column_letter(22 + ci)
        _col_tot  = get_column_letter(2  + ci)
        _r_tg_fl  = _r_tot2["flex"]; _r_tg_tot = _r_tot2[None]
        c = ws_mov.cell(row=_r_pflex, column=22+ci,
                        value=f"=IFERROR({_col_flex}{_r_tg_fl}/{_col_tot}{_r_tg_tot},\"\")")
        c.number_format = "0%"; c.font = _mfont()
        c.alignment = _maln("center"); c.border = _mall_bdr()

    # ════════════════════════════════════════════════════════════════
    # BLOQUE 3 — 200MBPS / >=400MBPS / TV COMPLETA
    # ════════════════════════════════════════════════════════════════
    _R3 = max(_r_tot2["reg"], _r_tot2["flex"]) + 7

    def _vel_pivot(filtro_fn, dias):
        _d = _altas_mes[filtro_fn(_altas_mes)].copy()
        if "ZONAL" in _d.columns and "zonal" not in _d.columns:
            _d = _d.rename(columns={"ZONAL": "zonal"})
        _p = _d.groupby([_d["_fa"].dt.normalize(), "zonal"])["Q"].sum().unstack(fill_value=0)
        return _p.reindex(index=dias, columns=_MOV_ZONALES, fill_value=0)

    _piv_200 = _vel_pivot(
        lambda d: d["velocidad_ba"].astype(str).str.strip() == "200", _dias_mes)
    _piv_400 = _vel_pivot(
        lambda d: pd.to_numeric(d["velocidad_ba"], errors="coerce").fillna(0) >= 400, _dias_mes)
    _piv_tv  = _vel_pivot(
        lambda d: d["sub_producto"].astype(str).str.upper().str.contains("COMPLETA TV", na=False),
        _dias_mes)

    _BLOQ3 = [
        (1,  "200MBPS",     _piv_200),
        (11, ">=400MBPS",   _piv_400),
        (21, "COMPLETA TV", _piv_tv),
    ]

    for _c_ini, _titulo, _piv in _BLOQ3:
        cell = ws_mov.cell(row=_R3, column=_c_ini, value=_titulo)
        cell.font = _mfont(bold=True); cell.alignment = _maln("left")

        # Encabezados: DIA + ZONALES + Total general
        _hdrs3 = ["DIA"] + _MOV_ZONALES + ["Total general"]
        for ci, hdr in enumerate(_hdrs3):
            cell = ws_mov.cell(row=_R3+1, column=_c_ini+ci, value=hdr)
            cell.fill      = _mfill(_MOV_B3_BG)
            cell.font      = _mfont(bold=True, color=_MOV_B3_FG)
            cell.alignment = _maln("center")
            cell.border    = _mall_bdr()

        for di, (dia, lbl) in enumerate(zip(_dias_mes, _dias_labels)):
            r = _R3 + 2 + di
            vals = [int(_piv.loc[dia, zon]) if zon in _piv.columns else 0
                    for zon in _MOV_ZONALES]
            for ci, v in enumerate([lbl] + vals + [sum(vals)]):
                cell = ws_mov.cell(row=r, column=_c_ini+ci, value=v)
                cell.font = _mfont(); cell.alignment = _maln("center" if ci > 0 else "left")
                cell.border = _mall_bdr()

        _r_tg3 = _R3 + 2 + len(_dias_mes)
        cell = ws_mov.cell(row=_r_tg3, column=_c_ini, value="Total general")
        cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
        for ci in range(_NZ + 1):
            _cl = get_column_letter(_c_ini + 1 + ci)
            c_t = ws_mov.cell(row=_r_tg3, column=_c_ini+1+ci,
                              value=f"=SUM({_cl}{_R3+2}:{_cl}{_r_tg3-1})")
            c_t.font = _mfont(bold=True); c_t.alignment = _maln("center"); c_t.border = _mall_bdr()

    # ════════════════════════════════════════════════════════════════
    # BLOQUE 4 — VENDEDORES C/ VENTA MES ACTUAL / MES ANTERIOR / RATIO
    # ════════════════════════════════════════════════════════════════
    _R4 = _R3 + 2 + len(_dias_mes) + 7

    # Pivot mes actual: count distinct DNI_VENDEDOR por día y zonal
    _piv_vdd_act = _pivot_dia_zon(_altas_mes,  "_fa", "DNI_VENDEDOR", "nunique", _dias_mes, _MOV_ZONALES)
    # Pivot mes anterior: misma cantidad de días (alinear al día del mes, no fecha absoluta)
    _piv_vdd_ant = _pivot_dia_zon(_altas_ant, "_fa", "DNI_VENDEDOR", "nunique", _dias_ant, _MOV_ZONALES)
    # Reindexar dias_ant con índice numérico de día (1..n) para alinear con dias_mes
    _vdd_ant_by_day = _piv_vdd_ant.copy()
    _vdd_ant_by_day.index = range(1, len(_dias_ant)+1)
    _vdd_act_by_day = _piv_vdd_act.copy()
    _vdd_act_by_day.index = range(1, len(_dias_mes)+1)

    _BLOQ4 = [
        (1,  "VENDEDORES C/ VENTA MES ACTUAL",    _vdd_act_by_day, _dias_labels, "DIA"),
        (11, "VENDEDORES C/ VENTA MES ANTERIOR",  _vdd_ant_by_day,
             [d.strftime("%a %d").lower().replace("mon","lun").replace("tue","mar")\
              .replace("wed","mié").replace("thu","jue").replace("fri","vie")\
              .replace("sat","sáb").replace("sun","dom") for d in _dias_ant], "Fecha"),
    ]

    _r_tg_vdd = {}
    for _c_ini, _titulo, _piv_v, _dlabels, _dheader in _BLOQ4:
        cell = ws_mov.cell(row=_R4, column=_c_ini, value=_titulo)
        cell.font = _mfont(bold=True); cell.alignment = _maln("left")

        _hdrs4 = [_dheader] + _MOV_ZONALES + ["Total general"]
        for ci, hdr in enumerate(_hdrs4):
            cell = ws_mov.cell(row=_R4+1, column=_c_ini+ci, value=hdr)
            cell.fill      = _mfill(_MOV_B4_BG)
            cell.font      = _mfont(bold=True, color=_MOV_B4_FG)
            cell.alignment = _maln("center")
            cell.border    = _mall_bdr()

        _n_dias_v = len(_dlabels)
        for di in range(_n_dias_v):
            r = _R4 + 2 + di
            _dia_idx = _piv_v.index[di] if di < len(_piv_v) else None
            if _dia_idx is not None:
                vals = [int(_piv_v.loc[_dia_idx, zon]) if zon in _piv_v.columns else 0
                        for zon in _MOV_ZONALES]
            else:
                vals = [0] * _NZ
            for ci, v in enumerate([_dlabels[di]] + vals + [sum(vals)]):
                cell = ws_mov.cell(row=r, column=_c_ini+ci, value=v)
                cell.font = _mfont(); cell.alignment = _maln("center" if ci > 0 else "left")
                cell.border = _mall_bdr()

        _r_tg4 = _R4 + 2 + _n_dias_v
        _r_tg_vdd[_c_ini] = _r_tg4
        cell = ws_mov.cell(row=_r_tg4, column=_c_ini, value="Total general")
        cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
        for ci in range(_NZ + 1):
            _cl = get_column_letter(_c_ini + 1 + ci)
            c_t = ws_mov.cell(row=_r_tg4, column=_c_ini+1+ci,
                              value=f"=SUM({_cl}{_R4+2}:{_cl}{_r_tg4-1})")
            c_t.font = _mfont(bold=True); c_t.alignment = _maln("center"); c_t.border = _mall_bdr()

    # Columna RATIO (U): VDD mes actual / VDD mes anterior, por zonal
    _r_ratio_start = _R4 + 1
    cell = ws_mov.cell(row=_R4, column=21, value="VENDEDORES C/ VENTA")
    cell.font = _mfont(bold=True); cell.alignment = _maln("left")
    _hdrs_ratio = [""] + _MOV_ZONALES + ["Total general"]
    for ci, hdr in enumerate(_hdrs_ratio):
        cell = ws_mov.cell(row=_r_ratio_start, column=21+ci, value=hdr)
        cell.fill      = _mfill(_MOV_B4_BG)
        cell.font      = _mfont(bold=True, color=_MOV_B4_FG)
        cell.alignment = _maln("center")
        cell.border    = _mall_bdr()

    _n_ratio = min(len(_dias_mes), len(_dias_ant))
    for di in range(_n_ratio):
        r        = _R4 + 2 + di
        r_act_row = _R4 + 2 + di   # misma fila relativa que el bloque actual
        r_ant_row = _R4 + 2 + di   # el bloque anterior tiene mismas filas
        for ci in range(_NZ + 1):
            _col_act = get_column_letter(2  + ci)
            _col_ant = get_column_letter(12 + ci)
            c = ws_mov.cell(row=r, column=22+ci,
                            value=f"=IFERROR({_col_act}{r_act_row}/{_col_ant}{r_ant_row},\"\")")
            c.number_format = "0%"; c.font = _mfont()
            c.alignment = _maln("center"); c.border = _mall_bdr()

    # Total ratio
    _r_tg_ratio = _R4 + 2 + _n_ratio
    cell = ws_mov.cell(row=_r_tg_ratio, column=21, value="Total general")
    cell.font = _mfont(bold=True); cell.alignment = _maln("center"); cell.border = _mall_bdr()
    _r_act_tg = _r_tg_vdd[1]; _r_ant_tg = _r_tg_vdd[11]
    for ci in range(_NZ + 1):
        _col_act = get_column_letter(2  + ci)
        _col_ant = get_column_letter(12 + ci)
        c = ws_mov.cell(row=_r_tg_ratio, column=22+ci,
                        value=f"=IFERROR({_col_act}{_r_act_tg}/{_col_ant}{_r_ant_tg},\"\")")
        c.number_format = "0%"; c.font = _mfont(bold=True)
        c.alignment = _maln("center"); c.border = _mall_bdr()

    # ── Ancho de columnas MOVISTAR ────────────────────────────────────────
    ws_mov.column_dimensions["A"].width = 9
    for _ci in range(2, 30):
        ws_mov.column_dimensions[get_column_letter(_ci)].width = 10

    # ── VDD3: hoja vacía — tablas dinámicas nativas creadas con xlwings ──
    # (se insertan al final del script, después de que openpyxl cierre el archivo)
    writer.book.create_sheet("VDD3")


    # ── Hojas de datos ───────────────────────────────────────
    rt_df.to_excel(writer, sheet_name="RT", index=False)
    altas_df.to_excel(writer, sheet_name="ALTAS", index=False)
    altas_ant_df.to_excel(writer, sheet_name="ALTAS_MES_PASADO", index=False)
    writer.sheets["ALTAS_MES_PASADO"].sheet_state = "hidden"

    # ── Formato RT y ALTAS ───────────────────────────────────
    # Columnas de fecha: aplican formato dd-mm-yy en ambas hojas
    # ALTAS: encabezados con colores en columnas específicas + TableStyleLight11
    # RT:    encabezados sin color + TableStyleLight8
    _DATE_COLS_NAMES = {"Fecha_de_alta", "Fecha_Registro", "Fecha_Venta", "Fecha_Alta"}

    # Columnas de ALTAS con fondo amarillo (FFFFFF00) y fuente roja (FF0000)
    _ALTAS_AMARILLO = {
        "PilotoPR+MONOBA", "DNI_INICIAL_VT", "Q", "TV", "CTRLNET",
        "segmento", "COM", "RIESG", "ESQUEMA", "VENDEDOR", "DNI", "SUPERVISOR",
    }
    # Columnas con fondo theme:3 azul claro (1F4E79 aprox) y fuente negra
    _ALTAS_AZUL = {"DNI_VENDEDOR", "Fecha_de_alta", "ORDEN", "FE"}
    # Columna peticion: theme:9 verde
    _ALTAS_VERDE = {"peticion"}

    def _fmt_sheet(ws, df_data, tbl_display_name, table_style, is_altas=False):
        nr = len(df_data)
        nc = len(df_data.columns)
        col_names = list(df_data.columns)

        # -- Solo encabezados (fila 1) y formato de número en columnas de fecha --
        # NO iteramos filas de datos para evitar corrupción de XML con tablas grandes
        for ci, col_name in enumerate(col_names, start=1):
            hcell = ws.cell(row=1, column=ci)
            hcell.alignment = Alignment(horizontal="center", vertical="center")

            if is_altas:
                if col_name in _ALTAS_AMARILLO:
                    hcell.fill = PatternFill("solid", fgColor="FFFF00")
                    hcell.font = Font(name="Aptos Narrow", size=10, bold=True, color="FF0000")
                elif col_name in _ALTAS_AZUL:
                    hcell.fill = PatternFill("solid", fgColor="1F4E79")
                    hcell.font = Font(name="Aptos Narrow", size=10, bold=True, color="000000")
                elif col_name in _ALTAS_VERDE:
                    hcell.fill = PatternFill("solid", fgColor="375623")
                    hcell.font = Font(name="Aptos Narrow", size=10, bold=True, color="000000")
                else:
                    hcell.font = Font(name="Aptos Narrow", size=10)
            else:
                # RT: fuente blanca en todos los encabezados
                hcell.font = Font(name="Aptos Narrow", size=10, color="FFFFFF")

        # Formato de fecha: usar column_dimensions para que aplique a toda la columna
        # sin tocar celda por celda
        for ci, col_name in enumerate(col_names, start=1):
            if col_name in _DATE_COLS_NAMES:
                # Aplicar solo en celdas de datos que ya fueron escritas por to_excel
                # Usamos iter_rows limitado a solo esa columna
                for row in ws.iter_rows(min_row=2, max_row=nr + 1,
                                        min_col=ci, max_col=ci):
                    for cell in row:
                        if cell.value is not None:
                            cell.number_format = "dd-mm-yy"

        # Tabla estructurada
        last_col = get_column_letter(nc)
        tbl = Table(displayName=tbl_display_name, ref=f"A1:{last_col}{nr + 1}")
        tbl.tableStyleInfo = TableStyleInfo(
            name=table_style,
            showFirstColumn=False, showLastColumn=False,
            showRowStripes=True,   showColumnStripes=False,
        )
        ws.add_table(tbl)

    _fmt_sheet(writer.sheets["RT"],    rt_df,    "Tbl_RT",    "TableStyleLight8",  is_altas=False)
    _fmt_sheet(writer.sheets["ALTAS"], altas_df, "Tbl_ALTAS", "TableStyleLight11", is_altas=True)
    rh.to_excel(writer, sheet_name="RH", index=False)
    ventory.to_excel(writer, sheet_name="VENTORY", index=False)

    # ── Hoja CON (oculta) — tabla de consultas DITO ──────────────
    df_con.to_excel(writer, sheet_name="CON", index=False)
    ws_con = writer.sheets["CON"]
    _con_rows = len(df_con)
    _con_cols = len(df_con.columns)
    _con_tbl = Table(
        displayName="CON",
        ref=f"A1:{get_column_letter(_con_cols)}{_con_rows + 1}"
    )
    _con_tbl.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight9",
        showFirstColumn=False, showLastColumn=False,
        showRowStripes=True,   showColumnStripes=False,
    )
    ws_con.add_table(_con_tbl)
    ws_con.sheet_state = "hidden"

    # ── Hojas MES y DIA ──────────────────────────────────────────
    # Estructura: DNI | Calificacion | Ventas | Instaladas
    # Ventas     = count de DNI_CNET en Tbl_RT   por Scoring
    # Instaladas = count de DNI_CNET en Tbl_ALTAS por Scoring
    # MES  → todo el período; DIA → solo día anterior (ayer_ts)

    def _build_mes_dia(rt_src, altas_src):
        """Construye el DataFrame DNI/Calificacion/Ventas/Instaladas.
        Cada DNI tiene siempre una fila REGULAR y una FLEX (cero si no hay datos)."""
        _SCORINGS = ["REGULAR", "FLEX"]

        # Todos los DNI presentes en cualquiera de las dos fuentes
        _todos_dni = pd.Series(
            pd.concat([rt_src["DNI_CNET"], altas_src["DNI_CNET"]]).dropna().unique(),
            name="DNI_CNET"
        )
        # Producto cartesiano DNI × [REGULAR, FLEX]
        _base = _todos_dni.to_frame().merge(
            pd.DataFrame({"Calificacion": _SCORINGS}), how="cross"
        )

        # RT: contar filas por DNI_CNET y Scoring
        _rt_grp = (
            rt_src.groupby(["DNI_CNET", "Scoring"])
            .size()
            .reset_index(name="Ventas")
            .rename(columns={"Scoring": "Calificacion"})
        )
        # ALTAS: contar filas por DNI_CNET y Scoring
        _al_grp = (
            altas_src.groupby(["DNI_CNET", "Scoring"])
            .size()
            .reset_index(name="Instaladas")
            .rename(columns={"Scoring": "Calificacion"})
        )

        _df = (_base
               .merge(_rt_grp, on=["DNI_CNET", "Calificacion"], how="left")
               .merge(_al_grp, on=["DNI_CNET", "Calificacion"], how="left"))
        _df["Ventas"]     = _df["Ventas"].fillna(0).astype(int)
        _df["Instaladas"] = _df["Instaladas"].fillna(0).astype(int)
        _df = _df.rename(columns={"DNI_CNET": "DNI"})
        return _df[["DNI", "Calificacion", "Ventas", "Instaladas"]].sort_values(["DNI", "Calificacion"]).reset_index(drop=True)

    # MES: fuentes completas del período
    _mes_rt    = rt_df[["DNI_CNET", "Scoring"]].copy()
    _mes_altas = altas_df[["DNI_CNET", "Scoring"]].copy()
    mes_df = _build_mes_dia(_mes_rt, _mes_altas)

    # DIA: solo registros de ayer
    _dia_rt    = rt_df[rt_df["Fecha_Registro"].dt.normalize() == ayer_ts][["DNI_CNET", "Scoring"]].copy()
    _dia_altas = altas_df[altas_df["Fecha_Alta"].dt.normalize() == ayer_ts][["DNI_CNET", "Scoring"]].copy()
    dia_df = _build_mes_dia(_dia_rt, _dia_altas)

    mes_df.to_excel(writer, sheet_name="MES", index=False)
    dia_df.to_excel(writer, sheet_name="DIA", index=False)
    writer.sheets["MES"].sheet_state = "hidden"
    writer.sheets["DIA"].sheet_state = "hidden"

    # ── Forzar fuente Aptos Narrow 11 en hojas sin formato propio ────
    # VDD1 ya tiene fuente y fills aplicados celda a celda — se excluye.
    # TDS ya fue procesada antes.
    for _ws_name in ["VDD2", "VDD3", "RT", "ALTAS", "RH", "VENTORY"]:
        if _ws_name in writer.sheets:
            _force_font_ws(writer.sheets[_ws_name])

print(f"Archivo generado (openpyxl): {ruta}")

# ════════════════════════════════════════════════════════════
# TABLAS DINÁMICAS — xlwings (requiere Excel instalado)
# Se abre el archivo ya guardado, se crean las dos PivotTables
# nativas en la hoja VDD3, apuntando a Tbl_VDD1 como origen,
# y se vuelve a guardar.
# ════════════════════════════════════════════════════════════
import xlwings as xw

# Constantes COM para PivotTable
_xlDatabase    = 1       # xlDatabase      — origen de datos rango/tabla
_xlRowField    = 1       # xlRowField
_xlColumnField = 2       # xlColumnField
_xlPageField   = 3       # xlPageField     — filtro de informe
_xlCount       = -4112   # xlCount
_xlPercentOfRow= 6       # xlPercentOfRow  — "% del total de filas"

def _crear_pivot(wb_api, sheet_vdd3_api, sheet_vdd1_api,
                 pivot_name, dest_cell_addr,
                 filter_field, filter_visible_items):
    """
    Crea una tabla dinámica nativa via COM.

    Parámetros
    ----------
    wb_api              : win32com Workbook
    sheet_vdd3_api      : win32com Worksheet destino (VDD3)
    sheet_vdd1_api      : win32com Worksheet origen  (VDD1)
    pivot_name          : str  nombre de la PivotTable
    dest_cell_addr      : str  celda esquina superior izquierda, ej. "A1"
    filter_field        : str  nombre del campo de filtro (ESQUEMA o ANTIG)
    filter_visible_items: list[str]  items que deben quedar visibles
    """
    # ── 1. Determinar rango de datos desde VDD1 (rango simple, sin tabla) ─
    # Buscamos primero una tabla Tbl_VDD1; si no existe usamos UsedRange.
    tbl_vdd1 = None
    for t in sheet_vdd1_api.ListObjects:
        if t.Name == "Tbl_VDD1":
            tbl_vdd1 = t
            break

    if tbl_vdd1 is not None:
        src_address = f"VDD1!{tbl_vdd1.Range.Address}"
    else:
        used = sheet_vdd1_api.UsedRange
        src_address = f"VDD1!{used.Address}"

    # ── 2. Crear PivotCache ───────────────────────────────────────
    pc = wb_api.PivotCaches().Create(
        SourceType=_xlDatabase,
        SourceData=src_address,
    )

    # ── 3. Crear PivotTable en la celda destino ───────────────────
    dest_cell = sheet_vdd3_api.Range(dest_cell_addr)
    pt = pc.CreatePivotTable(
        TableDestination=dest_cell,
        TableName=pivot_name,
    )

    # ── 4. Configurar campos ──────────────────────────────────────
    pt.ManualUpdate = True

    # Campo de FILTRO (página)
    pf_filter = pt.PivotFields(filter_field)
    pf_filter.Orientation = _xlPageField

    # Campo de FILAS: ZONAL
    pt.PivotFields("ZONAL").Orientation = _xlRowField
    pt.PivotFields("ZONAL").Position = 1

    # Campo de COLUMNAS: CLUSTER.ALTAS
    pt.PivotFields("CLUSTER.ALTAS").Orientation = _xlColumnField
    pt.PivotFields("CLUSTER.ALTAS").Position = 1

    # Campo de VALORES 1: Recuento → encabezado "[Q]"
    pt.AddDataField(
        pt.PivotFields("VENDEDOR"),
        "[Q]",
        _xlCount,
    )

    # Campo de VALORES 2: % del total de filas → encabezado "[%]"
    df2 = pt.AddDataField(
        pt.PivotFields("VENDEDOR"),
        "[%]",
        _xlCount,
    )
    df2.Calculation = _xlPercentOfRow
    df2.NumberFormat = "0%"

    # ── 5. Aplicar filtro de página (selección múltiple) ──────────
    pf_filter.EnableMultiplePageItems = True
    for item in pf_filter.PivotItems():
        item.Visible = (item.Name in filter_visible_items)

    # ── 6. Activar actualización y refrescar ─────────────────────
    pt.ManualUpdate = False
    pt.RefreshTable()

    return pt


def _color_semaforo(ws_api, cell_addr, valor):
    """Pinta el interior de una celda según semáforo de porcentaje."""
    # xlRgb: verde #63BE7B=6537339, amarillo #FFEB84=16772996, rojo #F8696B=16278891
    if valor is None or valor == "":
        return
    try:
        v = float(valor)
    except (TypeError, ValueError):
        return
    if v >= 0.95:
        rgb = 0x63BE7B   # verde
    elif v >= 0.90:
        rgb = 0xFFEB84   # amarillo
    else:
        rgb = 0xF8696B   # rojo
    # COM usa BGR: convertir RGB a BGR
    r = (rgb >> 16) & 0xFF
    g = (rgb >> 8)  & 0xFF
    b =  rgb        & 0xFF
    bgr = (b << 16) | (g << 8) | r
    ws_api.Range(cell_addr).Interior.Color = bgr


app = xw.App(visible=False, add_book=False)
try:
    wb      = app.books.open(str(ruta))
    wb_api  = wb.api
    ws_vdd3 = wb.sheets["VDD3"]
    ws_vdd1 = wb.sheets["VDD1"]
    ws_tds  = wb.sheets["TDS"]

    # ── Tabla dinámica 1 — ESQUEMA: PLANILLA + PART-TIME ─────────
    # Nota: _crear_pivot usa UsedRange de VDD1 directamente (Tbl_VDD1 no existe).
    # No se crea ListObject para evitar que Excel pise el formato de openpyxl.
    _crear_pivot(
        wb_api, ws_vdd3.api, ws_vdd1.api,
        pivot_name           = "TablaDinamica1",
        dest_cell_addr       = "A1",
        filter_field         = "ESQUEMA",
        filter_visible_items = ["PLANILLA", "PART-TIME"],
    )

    # ── Tabla dinámica 2 — ANTIG: <15d + >15d (fija en A16) ──────
    _crear_pivot(
        wb_api, ws_vdd3.api, ws_vdd1.api,
        pivot_name           = "TablaDinamica2",
        dest_cell_addr       = "A16",
        filter_field         = "ANTIG",
        filter_visible_items = ["<15d", ">15d"],
    )

    # ── VDD3: insertar 2 filas en blanco entre las dos tablas ────
    # Se insertan DESPUÉS de crear ambas tablas; la inserción en fila 14
    # empuja la segunda tabla hacia abajo sin afectar su estructura.
    ws_vdd3.api.Rows("14:15").Insert()

    # ── VDD3: fuente Aptos Narrow en todo el rango usado ─────────
    ws_vdd3.api.UsedRange.Font.Name = "Aptos Narrow"
    ws_vdd3.api.UsedRange.Font.Size = 11

    # ── Autoajustar filas en VDD3 ────────────────────────────────
    ws_vdd3.api.UsedRange.Rows.AutoFit()

    # ── VDD2: autoajuste de columnas ─────────────────────────────
    ws_vdd2 = wb.sheets["VDD2"]
    ws_vdd2.api.UsedRange.Columns.AutoFit()

    # ── Ocultar cuadrícula VDD3 y VDD2 via SheetViews ────────────
    # SheetViews(1) no requiere activate() ni ventana visible.
    # Envuelto en try/except para que un fallo no interrumpa el resto.
    try:
        ws_vdd3.api.SheetViews(1).DisplayGridlines = False
        ws_vdd2.api.SheetViews(1).DisplayGridlines = False
    except Exception as _e_grid:
        print(f"Aviso: no se pudo ocultar cuadrícula ({_e_grid})")

    # ── Semáforo N y AL en TDS (colores fijos, sin formato condicional) ─
    # Rangos: N6:N{r_tot} y AL6:AL{r_sup_last}
    # Los valores son porcentajes calculados; Excel los evalúa al abrir con xlwings.
    for _col_sem, _r_start, _r_end in [
        ("N",  _r0, _r_tot),
        ("AL", _r0, _r_sup_last),
    ]:
        for _row_i in range(_r_start, _r_end + 1):
            _cell_addr = f"{_col_sem}{_row_i}"
            _val = ws_tds.api.Range(_cell_addr).Value
            _color_semaforo(ws_tds.api, _cell_addr, _val)

    # ── Reordenar hojas: MOVISTAR, MiFibra, TDS, VDD1, VDD2, VDD3, resto ──
    _ORDEN_HOJAS = ["MOVISTAR", "MiFibra", "TDS", "VDD1", "VDD2", "VDD3"]
    _nombres_actuales = [s.name for s in wb.sheets]
    _pos = 0
    for _hn in _ORDEN_HOJAS:
        if _hn in _nombres_actuales:
            wb.sheets[_hn].api.Move(Before=wb.sheets[_pos].api)
            _pos += 1

    # ── Al abrir el libro debe quedar activa la hoja TDS ─────────
    wb.sheets["TDS"].api.Activate()

    wb.save()
    print("Tablas dinámicas creadas, hojas reordenadas, semáforo aplicado.")

    # ── Generar libro SEGUIMIENTO_VDD_FIJA con copia de VDD1/VDD2/VDD3 ──
    from datetime import date, timedelta
    _ayer = date.today() - timedelta(days=1)
    _fecha_str = _ayer.strftime("%d-%m-%Y")
    _ruta_seg  = _dir_salida / f"SEGUIMIENTO_VDD_FIJA_{_fecha_str}.xlsx"

    # El libro nuevo tiene 1 hoja vacía; se usa como ancla para el primer Copy
    wb_seg = app.books.add()

    for _nombre_hoja in ["VDD1", "VDD2", "VDD3"]:
        _sh_origen = wb.sheets[_nombre_hoja]
        # Copy con After= última hoja de wb_seg → la copia queda al final
        _sh_origen.api.Copy(After=wb_seg.sheets[-1].api)
        wb_seg.sheets[-1].name = _nombre_hoja

    # Eliminar la hoja vacía inicial que Excel creó al abrir el libro nuevo
    wb_seg.sheets[0].api.Delete()

    wb_seg.save(str(_ruta_seg))
    wb_seg.close()
    print(f"Seguimiento generado: {_ruta_seg}")

    # ── Generar libro AVANCE_VTAS_APPVENTORY con hojas MES y DIA ──
    _ruta_vtas = _dir_salida / f"AVANCE_VTAS_APPVENTORY_{ayer}.xlsx"
    wb_vtas = app.books.add()

    for _nombre_hoja in ["MES", "DIA"]:
        _sh_origen = wb.sheets[_nombre_hoja]
        _sh_origen.api.Copy(After=wb_vtas.sheets[-1].api)
        _copied = wb_vtas.sheets[-1]
        _copied.name = _nombre_hoja
        _copied.api.Visible = True

    wb_vtas.sheets[0].api.Delete()

    wb_vtas.save(str(_ruta_vtas))
    wb_vtas.close()
    print(f"Ventas AppVentory generado: {_ruta_vtas}")
finally:
    wb.close()
    app.quit()

print(f"Listo: {ruta}")