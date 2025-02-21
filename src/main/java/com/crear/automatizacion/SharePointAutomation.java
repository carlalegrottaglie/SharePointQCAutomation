package com.crear.automatizacion;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.concurrent.TimeUnit;

public class SharePointAutomation {
    public static WebDriver driver;

    public static void main(String[] args) throws IOException {
        // Configurar ChromeDriver
        System.setProperty("webdriver.chrome.driver", new File("./Drivers/chromedriver.exe").getCanonicalPath());

        driver = new ChromeDriver();
        driver.manage().window().maximize();
        WebDriverWait wait = new WebDriverWait(driver, java.time.Duration.ofSeconds(10));

        boolean primeraApertura = true;

        try (FileInputStream fileInputStream = new FileInputStream(new File("Proyectos.xlsx"));
             XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream)) {

            // Leer la hoja de "Proyectos"
            Sheet sheet = workbook.getSheet("Proyectos");

            // Comenzar desde la segunda fila (índice 1)
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;  // Saltar filas vacías

                // Leer la columna "Proyecto"
                String proyecto = row.getCell(0).getStringCellValue();
                // Leer la carpeta PDP (Columna 10)
                //String carpetaPDP = row.getCell(10).getStringCellValue();
                // Leer la carpeta de informes de avance (Columna 11)
                //String carpetaInformesAvance = row.getCell(11).getStringCellValue();
                // Leer la carpeta de riesgos (Columna 12)
               // String carpetaRiesgos = row.getCell(12).getStringCellValue();
                // Leer la carpeta de cronogramas (Columna 13)
               // String carpetaCronograma = row.getCell(13).getStringCellValue();
                LocalDate fechaInicio = row.getCell(4).getLocalDateTimeCellValue().toLocalDate();
                LocalDate fechaFin = row.getCell(5).getLocalDateTimeCellValue().toLocalDate();

                System.out.println("NOMBRE DE PROYECTO: " + proyecto);

                // Navegar y verificar archivos en las distintas carpetas
                if (primeraApertura) {
                    // Pausa para permitir el inicio de sesión manual
                    driver.get("https://aisrl.sharepoint.com/:f:/s/proyectos/EmpIKnwiF5BLhtKLS-GudOcBis9ZdFyeT_46-qlWzCh4Qw?e=tepLW2"); // Intento de abrir la carpeta PDP para que aparezca el login

                    System.out.println("Por favor, inicie sesión en la página SharePoint y presione Enter para continuar...");
                    System.in.read(); // Espera la entrada del usuario
                    primeraApertura = false; // Desactivar la pausa para las siguientes aperturas
                }

                // Verificar archivo PDP
                //driver.findElement(By.cssSelector("[name='Documentos']")).click();

                //scrollAndClick("[title='Proyectos Abiertos']");
                scrollAndClick("[title*='"+ proyecto + "']");


                verificarArchivoPDP(driver);
                driver.findElement(By.cssSelector("[title='Gestión del Proyecto']")).click();

                // Verificar informes de avance

                wait.until(
                        ExpectedConditions.visibilityOfElementLocated(By.cssSelector("[title='Comunicación']")));
                driver.findElement(By.cssSelector("[title='Comunicación']")).click();

                wait.until(
                        ExpectedConditions.visibilityOfElementLocated(By.cssSelector("[title='Informes de Avance']")));
                driver.findElement(By.cssSelector("[title='Informes de Avance']")).click();

                navegarYVerificarInformesDeAvance(driver, fechaInicio, fechaFin);
                // Verificar cronogramas
                driver.findElement(By.xpath("//div[contains(text(),'Gestión del Proyecto')]")).click();
                wait.until(
                        ExpectedConditions.visibilityOfElementLocated(By.cssSelector("[title='Cronograma']")));
                driver.findElement(By.cssSelector("[title='Cronograma']")).click();

                verificarCronogramas(driver, fechaInicio, fechaFin);

                // Verificar riesgos
                driver.findElement(By.xpath("//div[contains(text(),'Gestión del Proyecto')]")).click();   wait.until(
                        ExpectedConditions.visibilityOfElementLocated(By.cssSelector("[title='Riesgos']")));
                driver.findElement(By.cssSelector("[title='Riesgos']")).click();
                verificarRiesgos(driver, fechaInicio, fechaFin);
                driver.findElement(By.xpath("//div[contains(text(),'Proyectos Abiertos')]")).click();
            }

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            driver.quit();
        }
    }
    // Función para scrollear y hacer click en un elemento
    public static void scrollAndClick(String locator){

        WebElement element = driver.findElement(By.cssSelector(locator));
        ((JavascriptExecutor) driver).executeScript("arguments[0].scrollIntoView(true);",element);
        element.click();
    }
    // Función para verificar si existe un archivo que comience con "PDP"
    private static void verificarArchivoPDP(WebDriver driver) {
        try {
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
            List<WebElement> archivos = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("button[data-automationid='FieldRenderer-name']")));

            for (WebElement archivo : archivos) {
                String nombreArchivo = archivo.getText();
                if (nombreArchivo.startsWith("PDP")) {
                    System.out.println("Archivo encontrado con nombre que comienza con PDP: " + nombreArchivo);
                    return;  // Salir una vez encontrado
                }
            }

            System.out.println("No se encontró ningún archivo que comience con PDP.");

        } catch (Exception e) {
            System.out.println("Error al buscar archivo con nombre PDP: " + e.getMessage());
        }
    }

    private static void navegarYVerificarInformesDeAvance(WebDriver driver, LocalDate fechaInicio, LocalDate fechaFin) {
        try {
            // Esperar a que los archivos estén visibles
            WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(30));
            List<WebElement> archivos = wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector("button[data-automationid='FieldRenderer-name']")));

            Set<String> nombresArchivos = new HashSet<>();
            for (WebElement archivo : archivos) {
                String nombreArchivo = archivo.getText();
                nombresArchivos.add(nombreArchivo);
                System.out.println("Archivo encontrado: " + nombreArchivo);
            }

            // Verificar informes de avance
            verificarInformes(nombresArchivos, fechaInicio, fechaFin);

        } catch (Exception e) {
            System.out.println("Error al navegar o verificar informes de avance: " + e.getMessage());
        }
    }

    private static void verificarCronogramas(WebDriver driver, LocalDate fechaInicio, LocalDate fechaFin) {
        try {
            // Verificar versiones del cronograma
            verificarVersiones(driver, fechaInicio, fechaFin);

        } catch (Exception e) {
            System.out.println("Error al navegar o verificar cronogramas: " + e.getMessage());
        }
    }

    private static void verificarRiesgos(WebDriver driver, LocalDate fechaInicio, LocalDate fechaFin) {
        try {
            // Verificar versiones de riesgos
            verificarVersiones(driver, fechaInicio, fechaFin);

        } catch (Exception e) {
            System.out.println("Error al navegar o verificar riesgos: " + e.getMessage());
        }
    }

    private static void verificarInformes(Set<String> nombresArchivos, LocalDate fechaInicio, LocalDate fechaFin) {
        List<String> patronesInformesEsperados = new ArrayList<>();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM");

        LocalDate fechaActual = fechaInicio.plusMonths(1); // Iniciar desde el mes siguiente
        while (!fechaActual.isAfter(fechaFin)) {
            patronesInformesEsperados.add(fechaActual.format(formatter));
            fechaActual = fechaActual.plusMonths(1);
        }

        List<String> informesPresentesConFormatoCorrecto = new ArrayList<>();
        List<String> informesPresentesConFormatoIncorrecto = new ArrayList<>();
        List<String> informesFaltantes = new ArrayList<>(patronesInformesEsperados);

        for (String nombreArchivo : nombresArchivos) {
            boolean formatoCorrecto = nombreArchivo.startsWith("IA") &&
                    nombreArchivo.matches(".*\\d{4}-\\d{2}.*") &&
                    nombreArchivo.matches(".*v\\d+.*");

            // Extraer el mes y año del nombre del archivo
            String fechaMes = null;
            if (nombreArchivo.matches(".*\\d{4}-\\d{2}.*")) {
                fechaMes = nombreArchivo.replaceAll(".*(\\d{4}-\\d{2}).*", "$1");
            }

            if (fechaMes != null && patronesInformesEsperados.contains(fechaMes)) {
                if (formatoCorrecto) {
                    informesPresentesConFormatoCorrecto.add(nombreArchivo);
                } else {
                    informesPresentesConFormatoIncorrecto.add(nombreArchivo);
                }
                informesFaltantes.remove(fechaMes);
            }
        }

        // Mostrar resultados
        System.out.println("Informes de avance presentes con formato correcto: " + informesPresentesConFormatoCorrecto);
        System.out.println("Informes de avance presentes con formato incorrecto: " + informesPresentesConFormatoIncorrecto);
        System.out.println("Informes de avance faltantes: " + informesFaltantes);
    }

    private static void verificarVersiones(WebDriver driver, LocalDate fechaInicio, LocalDate fechaFin) {
        try {
            // Esperar y hacer clic en el botón de los tres puntos
            WebElement botonOpciones = new WebDriverWait(driver, Duration.ofSeconds(10)).until(
                    ExpectedConditions.elementToBeClickable(By.xpath("//button[contains(@class, 'ms-Button')]")));
            botonOpciones.click();

            // Seleccionar "Ver historial de versiones"
            WebElement verHistorial = new WebDriverWait(driver, Duration.ofSeconds(10)).until(
                    ExpectedConditions.elementToBeClickable(By.xpath("//span[text()='Historial de versiones']")));
            verHistorial.click();

            // Esperar a que se cargue el historial de versiones
            List<WebElement> versiones = new WebDriverWait(driver, Duration.ofSeconds(10)).until(
                    ExpectedConditions.visibilityOfAllElementsLocatedBy(By.cssSelector(".ms-ContextualMenu-link")));

            Set<String> mesesConVersiones = new HashSet<>();
            DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyy-MM");

            for (WebElement version : versiones) {
                String textoVersion = version.getText();
                if (textoVersion.matches(".*\\d{2}/\\d{2}/\\d{4}.*")) {
                    String fechaVersion = textoVersion.replaceAll(".*(\\d{2}/\\d{2}/\\d{4}).*", "$1");
                    LocalDate fecha = LocalDate.parse(fechaVersion, DateTimeFormatter.ofPattern("dd/MM/yyyy"));
                    if (!fecha.isBefore(fechaInicio) && !fecha.isAfter(fechaFin)) {
                        String mesVersion = fecha.format(formatter);
                        mesesConVersiones.add(mesVersion);
                    }
                }
            }

            // Generar lista de meses esperados
            Set<String> mesesEsperados = new HashSet<>();
            LocalDate fechaActual = fechaInicio.plusMonths(1);
            while (!fechaActual.isAfter(fechaFin)) {
                mesesEsperados.add(fechaActual.format(formatter));
                fechaActual = fechaActual.plusMonths(1);
            }

            // Determinar meses faltantes
            mesesEsperados.removeAll(mesesConVersiones);

            // Mostrar resultados
            System.out.println("Meses con versiones: " + mesesConVersiones);
            System.out.println("Meses sin versiones: " + mesesEsperados);

        } catch (Exception e) {
            System.out.println("Error al verificar el historial de versiones: " + e.getMessage());
        }
    }
}
