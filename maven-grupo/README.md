### Word Table Reader & Excel Writer 

1. Buscar archivo main.java
2.  buscar una variable que sea similar a esta 
`static String rutaAlumnosPK = "MATRICULA2022/FICHA INSCRIPCION PREKINDER 2022/PREKINDER 2022";`
3.  Copy/Paste todo el codigo anterior y dentro de las comillas " " ingresar una nueva ruta donde existan archivos de alumnos a procesar. 
	- (No debe ser la ruta del archivo del alumno!! debe ser la ruta de la carpeta que contiene los archivos ) 
4. Luego se puede cambiar el nombre de variable al que desees y que distinga el curso del alumno.

5. Buscar una variable similar a la siguiente. 
`	static List<String> listaAlumnosPK = listarArchivosDocx(rutaAlumnosPK);
`

6. Cambiar el nombre de la variable por uno que identifique al curso y que coincida con el curso de la ruta anterior. 

7.  Y finalmente, dentro de esta parte: `listarArchivosDOCX( AQUI VA LA VARIABLE DE LA LISTA)`  ingresar el nombre de la variable de lista donde se indicó.

8. Para terminar, copy/paste lo siguiente `procesoFinal(rutaAlumnos, listaAlumnos);`
y pegarlo dentro de los corchetes de public static void main. 

9. Hacer click derecho en main.java -> run as -> Java application 
- [x] Y listo! Generara los informes excel en la carpeta INFORME2022 o donde tu estimes conveniente. 






