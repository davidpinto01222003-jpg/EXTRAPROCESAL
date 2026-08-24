// Datos de los derechos de petición por prescripción de multas de tránsito

// Membrete (constante: es el papel del abogado, no del peticionario)
const MEMBRETE = {
  tel: "+57 317 866 6318",
  email: "22jorgecaicedo@gmail.com",
};

// ---------------------------------------------------------------- peticionarios
const RAFAEL = {
  nombre: "RAFAEL LÓPEZ GUTIÉRREZ",
  cc: "1.098.650.278",
  domicilio: "____________________",
  emailNotif: "RafaelJoshua2611@gmail.com",
  ciudadNotif: "____________________",
  direccion: "________________________________",
  cel: "____________",
  ciudadFirma: "Bucaramanga",
  fechaLarga: "agosto de 2026",
};

const BRAHAYAN = {
  nombre: "BRAHAYAN CAMILO ROSSO VACCA",
  cc: "1.095.834.547",
  domicilio: "____________________",
  // no se suministró correo propio: se usa el del membrete, como en el modelo de Aguachica
  emailNotif: "22jorgecaicedo@gmail.com",
  ciudadNotif: "____________________",
  direccion: "________________________________",
  cel: "____________",
  ciudadFirma: "Bucaramanga",
  fechaLarga: "agosto de 2026",
};

const DESTINOS = [
  {
    archivo: "DP-PRESCRIPCION-BUCARAMANGA",
    entidad: "DIRECCIÓN DE TRÁNSITO DE BUCARAMANGA – DTB",
    dependencia: "Grupo de Cobro Coactivo / Oficina Asesora Jurídica",
    ciudad: "Bucaramanga, Santander",
    nombreCorto: "Bucaramanga",
    pet: RAFAEL,
    comparendos: [
      {
        numero: "68001000000020138305",
        fechaCorta: "19/06/2018",
        fechaLarga: "19 de junio de 2018",
        hora: "07:57",
        lugar: "CARRERA 21E CALLE 115*87, Bucaramanga",
        codigo: "C24",
        infraccion: "Conducir motocicleta sin observar las normas establecidas en el presente código.",
        resolucion: "000345024",
        fechaResolucion: "3 de agosto de 2018",
        coactivo: "293153",
        fechaCoactivo: "2 de agosto de 2019",
        placa: "LDL27C",
      },
      {
        numero: "68001000000017050097",
        fechaCorta: "06/02/2018",
        fechaLarga: "6 de febrero de 2018",
        hora: "00:00",
        lugar: "CARRERA 16 CALLE 104, Bucaramanga",
        codigo: "D12",
        infraccion: "Conducir un vehículo que, sin la debida autorización, se destine a un servicio diferente de aquel para el cual tiene licencia de tránsito.",
        resolucion: "000331492",
        fechaResolucion: "22 de marzo de 2018",
        coactivo: "282407",
        fechaCoactivo: "25 de octubre de 2018",
        placa: null,
      },
      {
        numero: "68001000000034947321",
        fechaCorta: "24/01/2023",
        fechaLarga: "24 de enero de 2023",
        hora: "13:57",
        lugar: "CALLE 105 # 26 97, Bucaramanga",
        codigo: "C35",
        infraccion: "No realizar la revisión técnico-mecánica en el plazo legal establecido o cuando el vehículo no se encuentre en adecuadas condiciones técnico-mecánicas o de emisión de gases, aun cuando porte los certificados correspondientes.",
        resolucion: "2023517668",
        fechaResolucion: "8 de marzo de 2023",
        coactivo: "386270",
        fechaCoactivo: "25 de septiembre de 2023",
        placa: null,
      },
    ],
  },
  {
    archivo: "DP-PRESCRIPCION-CURITI",
    entidad: "SECRETARÍA DE TRÁNSITO Y TRANSPORTE DE CURITÍ",
    dependencia: "O quien haga sus veces como autoridad de tránsito – Cobro Coactivo",
    ciudad: "Curití, Santander",
    nombreCorto: "Curití",
    pet: RAFAEL,
    comparendos: [
      {
        numero: "99999999000003896745",
        fechaCorta: "03/11/2018",
        fechaLarga: "3 de noviembre de 2018",
        hora: "18:10",
        lugar: "VÍA SAN GIL – BUCARAMANGA, KM 11 + 400",
        codigo: "D12",
        infraccion: "Conducir un vehículo que, sin la debida autorización, se destine a un servicio diferente de aquel para el cual tiene licencia de tránsito.",
        resolucion: "253",
        fechaResolucion: "18 de diciembre de 2018",
        coactivo: null,
        fechaCoactivo: null,
        placa: null,
      },
    ],
  },
  {
    archivo: "DP-PRESCRIPCION-EL-PLAYON",
    entidad: "SECRETARÍA DE TRÁNSITO Y TRANSPORTE DE EL PLAYÓN",
    dependencia: "O quien haga sus veces como autoridad de tránsito – Cobro Coactivo",
    ciudad: "El Playón, Santander",
    nombreCorto: "El Playón",
    pet: RAFAEL,
    comparendos: [
      {
        numero: "99999999000004573988",
        fechaCorta: "27/10/2020",
        fechaLarga: "27 de octubre de 2020",
        hora: "06:10",
        lugar: "VÍA BUCARAMANGA – SAN ALBERTO, KM 35 + 00",
        codigo: "D01",
        infraccion: "Guiar un vehículo sin haber obtenido la licencia de conducción correspondiente.",
        resolucion: "31-2021",
        fechaResolucion: "25 de febrero de 2021",
        coactivo: "SH-543",
        fechaCoactivo: "16 de junio de 2022",
        placa: null,
      },
    ],
  },
  {
    archivo: "DP-PRESCRIPCION-FLORIDABLANCA",
    entidad: "DIRECCIÓN DE TRÁNSITO DE FLORIDABLANCA",
    dependencia: "O quien haga sus veces como autoridad de tránsito – Cobro Coactivo",
    ciudad: "Floridablanca, Santander",
    nombreCorto: "Floridablanca",
    pet: BRAHAYAN,
    comparendos: [
      {
        numero: "68276000000020098264",
        fechaCorta: "27/07/2021",
        fechaLarga: "27 de julio de 2021",
        hora: "17:40",
        lugar: "PARALELA ORIENTAL, Floridablanca",
        codigo: "C24",
        infraccion: "Conducir motocicleta sin observar las normas establecidas en el presente código.",
        resolucion: "2021-319",
        fechaResolucion: "9 de octubre de 2021",
        coactivo: "MG-2060",
        fechaCoactivo: "17 de mayo de 2023",
        placa: null,
      },
    ],
  },
];

module.exports = { MEMBRETE, DESTINOS };
