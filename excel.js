import System; //Importación de librerias 
import DoasCore.Spectra; //Importación de librerias 


var oXL = new ActiveXObject("Excel.Application"); // Se activa un nuevo archivo de excel  
oXL.Visible = true;
var oWB = oXL.Workbooks.Add();
var oSheet = oWB.ActiveSheet; 
oSheet.Cells(1, 1).Value = "Pixel"; // Se exportan los datos del numero de pixel 
oSheet.Cells(1, 2).Value = "Current"; // Se exportan los datos de intensidad
var i;
for(i = 0; i < Specbar.CurrentSpectrum.NChannel; i++) // Se crea un ciclo para cada dato
{
oSheet.Cells(2 + i, 1).Value = i;
oSheet.Cells(2 + i, 2).Value = Specbar.CurrentSpectrum.Intensity[i];
Console.WriteLine("Put pixel " + i + " to Excel sheet");
}
Console.WriteLine("Exportación terminada"); // Se termina la exportación cuando el mensaje aparece en la salida del código.
