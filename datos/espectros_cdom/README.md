# Espectros de absorción — ap / ad / aphy / aCDOM

Archivos Excel de espectros de absorción óptica de agua (bio-óptica marina/limnológica),
guardados el 2026-08-06 a pedido del usuario. **No tienen relación con los scripts
jurídicos del resto del repositorio**; están aquí sólo como archivo de datos.

## Archivos

| Archivo | Campaña / solicitud | Nº estaciones | Códigos de muestra |
|---|---|---|---|
| `Espectros_apadaphyCDOM_Sol000822_1.xlsx` | Sol 0008-22 | 23 | `E1`…`E31` |
| `Espectros_apadaphyCDOM_Sol001321.xlsx` | Sol 0013-21 | 32 | `SC-21-0101`…`SC-21-0231` |
| `Espectros_apadaphyCDOM_Sol002423_1.xlsx` | Sol 0024-23 | 15 | `SC-23-0513`…`SC-23-0548` (encabezados de hojas de datos usan `E1`…`E31`) |
| `Espectros_apadaphyCDOM_Sol003422_1.xlsx` | Sol 0034-22 | 29 | `SC-22-0415`…`SC-22-0479` |

Nota: el usuario envió `Sol000822_1` y `Sol000822_1_1`, que son **byte a byte idénticos**
(md5 `c74d473ac240a8a6e943c86446a3897c`). Se conservó una sola copia.

## Estructura común (7 hojas por libro)

| Hoja | Contenido | Rango espectral | Filas |
|---|---|---|---|
| `ap` | absorción del particulado total [m⁻¹] | 800 → 400 nm (1 nm) | 401 |
| `ad` | absorción del detritus / no-algal [m⁻¹] | 800 → 400 nm | 401 |
| `aphy` | absorción del fitoplancton (`ap − ad`) [m⁻¹] | 800 → 400 nm | 401 |
| `aCDOM` | absorción de materia orgánica disuelta coloreada [m⁻¹]; primero el bloque **medido** y a la derecha un segundo bloque **modelado** (`MODELADO` / `ESPECTROS MODELADOS`) con su propia columna `Wavelength (nm)` | 800 → 250 nm | 551 |
| `ax(443)` | resumen a 443 nm por estación: `ap(443)`, `ad(443)`, `aphy(443)`, `aCDOM(443)`, `aCDOM(443)-Modelado`, `TOTAL-CDOM`, `TOTAL-CDOM Modelado`, más bloques de porcentajes/relaciones | — | ~18–35 |
| `aphy-Especifico` | aphy específico (normalizado por clorofila) [m² mg⁻¹] | 800 → 400 nm | 401 |
| `Varias longitudes de onda` | valores extraídos en 715, 676, 650, 630, 555, 510, 488, 443, 440 y 412 nm para `aCDOM`, `ad`, `aphy` (y `ap` en algunos libros); además razones azul/rojo `B/R 440-676` y `B/R 443-677` con los cortes de tamaño: microfitoplancton `< 2.5`, nanofitoplancton `2.5–3.0`, picofitoplancton `> 3.0` | — | 42–55 |

Layout de todas las hojas espectrales: columna A = `Wavelength (nm)` en orden descendente,
una columna por estación con el código de muestra en la fila 1.
