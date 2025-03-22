using Google.Apis.Sheets.v4;
using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using System.Security.Cryptography.X509Certificates;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using System;
using Google.Apis.Sheets.v4.Data;

public class GoogleSheetsAPI : MonoBehaviour
{

    [Header("Google Sheets Information")]           // encabezado que aparece en sima de las variables en el editor de unity
    [SerializeField] private string spreadSheetID;  // ID del documento de google sheets
    [SerializeField] private string sheetID;        // ID de la hoja que se está editando

    [Header("Data from GoogleSheets")]                  //encabezao que aparece en sima de las variables en el editor de unity
    [SerializeField] private string getDataInRange;     // el rango de celdas de la que se obtendran los datos


    //es el correo que da el servisio de google sheets para poder editarlo desde unity
    private string serviceAccountEmail = "googlesheetunity@unityapi-454423.iam.gserviceaccount.com";
    //es el nombre del archivo con el certificado
    private string certificateName = "unityapi-454423-b92c42b92a4e.p12";
    //es la ruta donde se encuentra el sertificado
    private string certificatePath;

    //es el servicio de google sheets que creamos
    private static SheetsService googleSheetService;

    //se crea una clase de filas
    [Serializable]
    public class Row
    {
        //una lista de celdas
        public List<string> cellData = new List<string>();
    }

    //se crea una clase que almacene las listas de celdas
    [Serializable]
    public class RowList
    {
        //lista de filas
        public List<Row> rows = new List<Row>();
    }

    //se crea e inicializa un objeto de tipo RowList
    public RowList DataFromGoogleSheets = new RowList();

    [Header("Write Data from Unity")]                       //encabezao que aparece en sima de las variables en el editor de unity
    [SerializeField] private string writeDataInRange;       //rango de celdas en las que se escribiran datos

    public RowList WriteDataFromUnity = new RowList();      //estructura de datos

    [Header("Delete Data In GoogleSheets")]                 //encabezao que aparece en sima de las variables en el editor de unity
    [SerializeField] private string deleteDataInRange;      //esto es para definir el rango desde el inspector

    void Start()
    {
        //definimos la ruta del certificado
        certificatePath = "/StreamingAssets/" + certificateName;

        //importamos el certificado, lo que está entre comillas es la contraseña que se nos da en la pagina del servisio de googlesheets
        var certificate = new X509Certificate2(Application.dataPath + certificatePath, "notasecret", X509KeyStorageFlags.Exportable);

        //enviar el certificado a google sheets para tener el acceso
        ServiceAccountCredential credential = new ServiceAccountCredential(
            new ServiceAccountCredential.Initializer(serviceAccountEmail)
            {
                Scopes = new[] { SheetsService.Scope.Spreadsheets }
            }.FromCertificate(certificate));


        //terminar de inicializar el certificado de googlesheets
        googleSheetService = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "GoogleSheets API for Unity"
        });
        
        //ejecusion de las 3 opciones de funciones (salio verso sin hacer esfuerso)
        //ReadData();
        //WriteData();
        //DeleteData();
    }

    //funcion que permite leer la informacion de GoogleSheets
    public void ReadData()
    {
        //rango de celdas que se leeran en googlesheets
        string range = sheetID + "!" + getDataInRange;

        //se crea un requerimiento 
        var request = googleSheetService.Spreadsheets.Values.Get(spreadSheetID, range);
        //se crea una respuesta
        var reponse = request.Execute();
        //guarda los valores obtenidos
        var values = reponse.Values;

        //verifica si los valores no son nulos
        if (values != null && values.Count > 0)
        {
            //por cada fila de datos...
            foreach ( var row in values)
            {
                //se crea una nueva fila
                Row newRow = new Row();
                //y se agrega a la lista de filas
                DataFromGoogleSheets.rows.Add(newRow);
                //por cada datos que está en cada fila...
                foreach (var data in row)
                {
                    //se agrega en la lista de celdas
                    newRow.cellData.Add(data.ToString());
                }
            }
        }

    }

    //funcion en la que se escribiran los datos en la hoja de google
    public void WriteData()
    {
        string range = sheetID + "!" + writeDataInRange;        //se define el rango de celdas que seran escritas

        var valueRange = new ValueRange();                      //rango de valores
        var cellData = new List<object>();                      //lista de celdas
        var arrows = new List<IList<object>>();                 //lista bde filas
        //ordena los datos
        foreach (var row in WriteDataFromUnity.rows)            
        {
            cellData = new List<object>();                      //se crea una lista de celdas
            foreach(var data in row.cellData)                   //se agrega el dato de cada celda
            {
                cellData.Add(data);
            }
            arrows.Add(cellData);                               //se agrega la fila
        }

        valueRange.Values = arrows;

        var request = googleSheetService.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        var reponse = request.Execute();
    }

    //funsion para eliminar datos de la hoja de google
    public void DeleteData()
    {
        string range = sheetID + "!" + deleteDataInRange;       //rango de celdas a las que se quiere eliminar la informacion

        //se crea una variable con la que se liminaran los datos dentro del rango de celdas
        var deleteData = googleSheetService.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadSheetID, range);
        deleteData.Execute();                                   //se ejecuta la variable para que esta elimine los datos
    }

}
