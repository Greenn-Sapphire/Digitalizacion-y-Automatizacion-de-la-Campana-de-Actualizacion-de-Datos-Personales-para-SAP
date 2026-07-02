function separarDireccion(Direccion) {
  //Direccion = "SM 222, MZ 2, Lote 5 José María Pino Suárez, 20, SM 222, Cancún, Benito Juarez, QR, MX, 77517"
  if (!Direccion) return [];

  // Separar por comas SIN eliminar vacíos
  const partes = Direccion.split(",").map(p => p.trim());

  // Calle: primeros 3 segmentos
  const Calle = partes.slice(0, 4).join(", ");

  // Campos restantes (respetando posiciones)
  const Num_Casa = partes[4] ?? "";
  const Colonia = partes[5] ?? "";
  const Poblacion = partes[6] ?? "";
  let Municipio = partes[7] ?? "";
  const Region = partes[8] ?? "";
  const Pais = partes[9] ?? "";
  const Codigo_Postal = partes[10] ?? "";

  if(Municipio){
    switch (Municipio) {
      case "Isla Mujeres":
        Municipio = 3;
        break;
      case "Benito Juarez":
        Municipio = 5;
        break;
      case "Solidaridad":
        Municipio = 8;
        break;
      case "Puerto Morelos":
        Municipio = 11;
        break;
      case "Tizimin":
        Municipio = 96;
        break;
      case "Valladolid":
        Municipio = 102;
        break;
      case "Merida":
        Municipio = 50;
        break;
      default:
        Municipio = 1010;
    }
  }

  console.log([Calle, Num_Casa, Colonia, Poblacion, Municipio, Region, Pais, Codigo_Postal])
  return [Calle, Num_Casa, Colonia, Poblacion, Municipio, Region, Pais, Codigo_Postal];
}
