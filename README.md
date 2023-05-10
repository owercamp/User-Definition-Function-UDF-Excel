
# Funciones personalizadas Soandes

![Project](https://img.shields.io/badge/Project-SOANDES-rgb(255,255,255)?labelColor=rgba(9,95,154,255)&logo=simple-icons&logoColor=rgba(9,120,154,255)) ![VBA](https://img.shields.io/badge/VBA-FUNCIONES%20DEFINIDAS%20POR%20USUARIO-rgb(25,63,102)?labelColor=rgb(37,150,190)&logo=OpenProject&logoColor=rgba(9,120,154,255)) ![VBA Application](https://img.shields.io/badge/VBA%20Application-EXCEL-rgb(25,63,102)?labelColor=rgb(0,0,0)&logo=World-Health-Organization&logoColor=rgba(9,120,154,255)) ![GitHub code size in bytes](https://img.shields.io/github/languages/code-size/owercamp/User-Definition-Function-UDF-Excel) ![GitHub repo file count](https://img.shields.io/github/directory-file-count/owercamp/User-Definition-Function-UDF-Excel) ![GitHub top language](https://img.shields.io/github/languages/top/owercamp/User-Definition-Function-UDF-Excel?color=yellowgreen) ![GitHub language count](https://img.shields.io/github/languages/count/owercamp/User-Definition-Function-UDF-Excel?color=orange) ![GitHub Sponsors](https://img.shields.io/github/sponsors/owercamp) ![GitHub issues](https://img.shields.io/github/issues/owercamp/User-Definition-Function-UDF-Excel) ![GitHub closed issues](https://img.shields.io/github/issues-closed/owercamp/User-Definition-Function-UDF-Excel) ![GitHub](https://img.shields.io/github/license/owercamp/User-Definition-Function-UDF-Excel) ![GitHub package.json version](https://img.shields.io/github/package-json/v/owercamp/User-Definition-Function-UDF-Excel) 


Este libro de Excel contiene un conjunto de funciones personalizadas desarrolladas en VBA para procesar datos que no son parte de las funciones nativas de Excel. A continuación se describen las funciones disponibles y cómo utilizarlas.

Funciones disponibles

___
#### IMEDICALFACTURE

Se busca obtener el valor a pagar en Avancys, el cual se encuentra en un libro de Excel proporcionado por el área de facturación. Para encontrar el valor correspondiente, se buscará el número de identificación en una columna vertical y el código CUPS en una fila horizontal, y se tomará el valor de la celda donde se cruzan ambas. Se utilizarán las funciones de búsqueda y referencia de celdas de Excel para llevar a cabo esta tarea.


##### Parámetros de entrada:


- identity: [integer] Número de identificación a buscar.
- rng_identity: [Range] Matriz del rango en donde se consultara los numeros de identificación.
- cups: [string] Código CUPS a buscar.
- rng_cups: [Range] Matriz del rango en donde se consultara los CUPS.

##### Valor de retorno:

[LongPtr] Es el valor a pagar por el servicio realizado segun la identificación del usuario y el examen de codigo cups.

##### Ejemplo de uso:

`
=IMEDICALFACTURE(identity, rng_identity, cups, rng_cups)
`
___

#### INTERPRETACION

La función UDF evalúa los resultados de un examen y determina si los valores de un índice están dentro del rango de referencia. Si el valor está dentro del rango, se clasifica como normal, de lo contrario se clasifica como anormal.

#### Parámetros de entrada:

- valorBuscado: [integer] es el valor del resultado a consultar dentro del rango de referencia.
- valorRango: [string] es el valor del rango de referencia representado 1000 - 5000.
- separador: [Variant] es el valor por el que se divide el rango de referencia.

##### Valor de retorno:

[String] Es el NORMAL o ANORMAL dependiendo si se encuentra dentro o fuera de los valores de referencia.

##### Ejemplo de uso:

`
=INTERPRETACION(valorBuscado, valorRango, separador)
`
___

#### BUSCAROP

La función BUSCAROP es una UDF (User-Defined Function) en Excel que se utiliza para buscar un valor dentro de una matriz de búsqueda y desplazarse hacia la izquierda o la derecha según la posición dada.

#### Parámetros de entrada:

- valor_buscado: [Variant] Es el valor que se desea buscar dentro de la matriz de búsqueda.
- rango_busqueda: [Range] Es el rango de celdas donde se buscará el valor.
- posicion: [Integer] Es el número de celdas que se desplazará la función hacia la izquierda (-) o hacia la derecha (+) desde la celda donde se encuentra el valor encontrado en la matriz.

##### Valor de retorno:

[Variant] valor en la matriz luego de desplazarse hacia la izquierda o hacia la derecha desde la posición encontrada en la búsqueda.

##### Ejemplo de uso:

`
=BUSCAROP(valor_buscado, rango_busqueda, posicion)
`
___

#### CONTARDATO

La función CONTARDATO en Excel es una User Defined Function (UDF) que permite contar la cantidad de veces que un valor determinado, ya sea un texto o número, aparece en un rango seleccionado y es visible.

#### Parámetros de entrada:

- data: [Range] Es el rango a verificar.
- text: [Variant] Es el valor que se desea buscar y contar si es visible o no.

##### Valor de retorno:

[Integer] Es un número entero que se devuelve como resultado luego de contar la cantidad de veces que el valor se encuentra en el rango seleccionado.

##### Ejemplo de uso:

`
=CONTARDATO(data, text)
`
___

#### FRAMINGHAM

La función FRAMINGHAM en Excel es una User Defined Function (UDF) que utiliza el modelo de Framingham para estimar el riesgo de enfermedad cardiovascular en una persona en función de varios factores de riesgo.

#### Parámetros de entrada:

- Age: [Integer] Edad de la persona (en años).
- Cholestrol: [Integer] Colesterol total de la persona (en mg/dL).
- Hdl: [Integer] Lipoproteína de alta densidad (HDL) de la persona (en mg/dL).
- Ts_tbs: [String] relación entre el colesterol total y el HDL (en formato "X/Y").
- Smoking: [String] indica si la persona fuma ("Fuma" si es fumador, de lo contrario "").
- Diabetes: [String] indica si la persona tiene diabetes ("Si" si tiene diabetes, de lo contrario "").
- Sex: [String] género de la persona ("Femenino" o "Masculino").

##### Valor de retorno:

[String] Es un cadena que indica el nivel de riesgo cardiovascular expresado como un porcentaje y una categoría de riesgo ("BAJO", "MODERADO", "ALTO" o "MUY ALTO").

##### Ejemplo de uso:

`
=FRAMINGHAM(Age, Cholestrol, Hdl, Ts_tbs, Smoking, Diabetes, Sex)
`
___

:+1: excelente
