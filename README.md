# reporting-excel

[![Visual Studio Team services](https://img.shields.io/vso/build/huchim/b3010c88-dae9-415b-a73f-d0d3d47d8073/3.svg)]()
[![NuGet Pre Release](https://img.shields.io/nuget/vpre/Jaguar.Reporting.Excel.svg?style=flat-square)]()

Permite manejar los reportes en formato de Excel (xlsx)

```csharp
// Crear el manejador de reportes.
var reportManager = new ReportManager(this.connection);

// Registrar este generador.
reportManager.AddGenerator(new ExcelGenerator());

// Agregar las variables de la consulta (opcional)
foreach (var k in this.Request.Query)
{
    reportManager.Variables.Add($"args.{k.Key}", k.Value);
}

// Abre el reporte (ver ReportRepository)
reportManager.Open(this.reportRepository.GetReport(id));

// Obtiene el resultado del reporte.
var results = reportManager.Process(type);

if (results == null)
{
    return this.NotFound("No hay reporte disponible.");
}

// Devolver el resultado al navegador.
return this.File(results.Data, results.MimeType, $"{id}{results.FileExtension}");
```

