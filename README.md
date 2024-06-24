# EXCEL_ACCESS_VS
Solcion para volcar una tabla excel en una tabla de access a traves de visual stucio .net VB
# EXCEL_ACCESS_VS

## 1. Ejecución de la Solución

### 1.1. Requisitos Previos
- .NET Framework 4.7.2 o superior.
- Microsoft Access.
- Microsoft Excel.
- Sistema operativo Windows.

### 1.2. Pasos de Instalación

1. **Clonar el Repositorio**
   - Clona el repositorio de GitHub a tu máquina local utilizando el comando:
     ```
     git clone https://github.com/GAB1972/EXCEL_ACCESS_VS.git
     ```

2. **Abrir la Solución en Visual Studio**
   - Navega a la carpeta del proyecto clonado:
     ```
     cd EXCEL_ACCESS_VS
     ```
   - Abre el archivo de la solución `EXCEL_ACCESS_VS.sln` con Visual Studio.

3. **Compilar la Solución**
   - En Visual Studio, compila la solución presionando `Ctrl+Shift+B` o seleccionando `Build > Build Solution` en el menú.

4. **Ejecutar la Aplicación**
   - En Visual Studio, presiona `F5` o selecciona `Debug > Start Debugging` para ejecutar la aplicación.

### 1.3. Configuración Inicial
- Al iniciar la aplicación por primera vez, se le pedirá que seleccione un archivo de base de datos Access (.mdb).
- En la carpeta raíz de la solución `EXCEL_ACCESS_VS\` se encuentra un fichero Access `.mdb` y un Excel con el que se han hecho las pruebas.
- Asegúrese de tener el archivo de base de datos en una ubicación accesible y seleccionarlo cuando se le solicite.

### 1.4. Uso de la Aplicación

1. **Importar Datos desde Excel**
   - Hacer clic en el botón `Importar desde Excel`, seleccionar el archivo `.xlsx` y confirmar la importación.
   - Se le notificará que los datos existentes serán sobrescritos.

2. **Ver y Editar Datos**
   - Hacer clic en el botón `Leer Datos` para cargar los datos en el `DataGridView`.
   - Editar los datos directamente en el `DataGridView`.

3. **Guardar Cambios**
   - Los cambios en el `DataGridView` se guardan automáticamente en la base de datos al finalizar la edición de una celda.

4. **Borrar Registro**
   - Para eliminar un registro, seleccione la fila correspondiente en el `DataGridView`.
   - Haga clic en el botón `Borrar Registro` para eliminar el registro seleccionado de la base de datos.

## 2. Arquitectura

### 2.1. Descripción General
La aplicación se compone de los siguientes componentes principales:
- **Interfaz de Usuario (UI)**: Implementada con Windows Forms para la interacción del usuario.
- **Gestión de Datos**: Módulos para la conexión y manipulación de la base de datos Access.
- **Módulo de Excel**: Para la lectura y procesamiento de archivos Excel.

### 2.2. Componentes
1. **FormMain**: Formulario principal de la aplicación que permite seleccionar archivos y mostrar datos.
2. **DatabaseManager**: Clase responsable de manejar la conexión y operaciones con la base de datos Access.
3. **ExcelManager**: Módulo para la lectura de archivos Excel y la conversión de datos en un formato adecuado para la base de datos.
4. **TableManager**: Módulo para crear y gestionar tablas en la base de datos Access.

## 3. Diseño

### 3.1. Interfaz de Usuario
El formulario principal (FormMain) contiene:
- Un botón para importar datos desde Excel (`ButtonImportExcel`).
- Un botón para leer y mostrar datos de la base de datos (`ButtonReadData`).
- Un `DataGridView` para mostrar los datos importados y permitir su edición.
- Etiquetas de estado (`LabelStatus`) para mostrar mensajes al usuario.

### 3.2. Funcionalidades
- **Importación de Excel**: Permite al usuario seleccionar un archivo Excel y cargar los datos en una base de datos Access.
- **Visualización y Edición de Datos**: Muestra los datos en un `DataGridView` y permite su edición.
- **Sincronización con Base de Datos**: Los cambios realizados en el `DataGridView` se reflejan en la base de datos Access.

### 3.3. Procedimiento
1. Al iniciar la aplicación, se pide al usuario que seleccione una base de datos Access.
2. El usuario puede importar datos desde un archivo Excel.
3. Los datos se muestran en un `DataGridView` donde se pueden editar.
4. Los cambios en el `DataGridView` se actualizan automáticamente en la base de datos Access.

## 4. Procedimiento de Instalación

### 4.1. Requisitos Previos
- .NET Framework 4.7.2 o superior.
- Microsoft Access.
- Microsoft Excel.
- Sistema operativo Windows.

### 4.2. Pasos de Instalación
1. **Clonar el Repositorio**
   - Clona el repositorio de GitHub a tu máquina local utilizando el comando:
     ```
     git clone https://github.com/GAB1972/EXCEL_ACCESS_VS.git
     ```

2. **Abrir la Solución en Visual Studio**
   - Navega a la carpeta del proyecto clonado:
     ```
     cd EXCEL_ACCESS_VS
     ```
   - Abre el archivo de la solución `EXCEL_ACCESS_VS.sln` con Visual Studio.

3. **Compilar la Solución**
   - En Visual Studio, compila la solución presionando `Ctrl+Shift+B` o seleccionando `Build > Build Solution` en el menú.

4. **Ejecutar la Aplicación**
   - En Visual Studio, presiona `F5` o selecciona `Debug > Start Debugging` para ejecutar la aplicación.

### 4.3. Configuración Inicial
- Al iniciar la aplicación por primera vez, se le pedirá que seleccione un archivo de base de datos Access (.mdb).
- En la carpeta raíz de la solución `EXCEL_ACCESS_VS\` se encuentra un fichero Access `.mdb` y un Excel con el que se han hecho las pruebas.
- Asegúrese de tener el archivo de base de datos en una ubicación accesible y seleccionarlo cuando se le solicite.

### 4.4. Uso de la Aplicación

1. **Importar Datos desde Excel**
   - Hacer clic en el botón `Importar desde Excel`, seleccionar el archivo `.xlsx` y confirmar la importación.
   - Se le notificará que los datos existentes serán sobrescritos.

2. **Ver y Editar Datos**
   - Hacer clic en el botón `Leer Datos` para cargar los datos en el `DataGridView`.
   - Editar los datos directamente en el `DataGridView`.

3. **Guardar Cambios**
   - Los cambios en el `DataGridView` se guardan automáticamente en la base de datos al finalizar la edición de una celda.

4. **Borrar Registro**
   - Para eliminar un registro, seleccione la fila correspondiente en el `DataGridView`.
   - Haga clic en el botón `Borrar Registro` para eliminar el registro seleccionado de la base de datos.

## 5. Ficheros de Especificación de los Generadores Utilizados para Prueba

### 5.1. Fichero Access
- **Nombre del Fichero**: `mibasededatos.mdb`
- **Ubicación**: Ruta raíz de la solución enviada.
- **Descripción**: Es un archivo de base de datos Access en blanco utilizado para almacenar los datos importados desde el archivo Excel.

### 5.2. Fichero Excel
- **Nombre del Fichero**: `miexcel.xlsx`
- **Ubicación**: Ruta raíz de la solución enviada.
- **Descripción**: Es un archivo Excel que contiene los datos de prueba para ser importados a la base de datos Access.
- **Contenido del Fichero Excel**:

