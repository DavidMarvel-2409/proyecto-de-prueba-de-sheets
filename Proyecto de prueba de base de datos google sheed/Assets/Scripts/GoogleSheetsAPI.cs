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
    [Header("Google Sheets Information")]
    [SerializeField] private string spreadSheetID;
    [SerializeField] private string sheetID;

    [Header("Data from GoogleSheets")]
    [SerializeField] private string getDataInRange;


    private string serviceAccountEmail = "googlesheetunity@unityapi-454423.iam.gserviceaccount.com";
    private string certificateName = "unityapi-454423-b92c42b92a4e.p12";
    private string certificatePath;

    private static SheetsService googleSheetService;
    [Serializable]
    public class Row
    {
        public List<string> cellData = new List<string>();
    }
    [Serializable]
    public class RowList
    {
        public List<Row> rows = new List<Row>();
    }

    public RowList DataFromGoogleSheets = new RowList();

    [Header("Write Data from Unity")]
    [SerializeField] private string writeDataInRange;

    public RowList WriteDataFromUnity = new RowList();

    [Header("Delete Data In GoogleSheets")]
    [SerializeField] private string deleteDataInRange; 

    void Start()
    {
        certificatePath = "/StreamingAssets/" + certificateName;

        var certificate = new X509Certificate2(Application.dataPath + certificatePath, "notasecret", X509KeyStorageFlags.Exportable);

        ServiceAccountCredential credential = new ServiceAccountCredential(
            new ServiceAccountCredential.Initializer(serviceAccountEmail)
            {
                Scopes = new[] { SheetsService.Scope.Spreadsheets }
            }.FromCertificate(certificate));


        googleSheetService = new SheetsService(new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential,
            ApplicationName = "GoogleSheets API for Unity"
        });
        //ReadData();
        //WriteData();
        DeleteData();
    }


    public void ReadData()
    {
        string range = sheetID + "!" + getDataInRange;

        var request = googleSheetService.Spreadsheets.Values.Get(spreadSheetID, range);
        var reponse = request.Execute();
        var values = reponse.Values;

        if (values != null && values.Count > 0)
        {
            foreach ( var row in values)
            {
                Row newRow = new Row();
                DataFromGoogleSheets.rows.Add(newRow);
                foreach (var data in row)
                {
                    newRow.cellData.Add(data.ToString());
                }
            }
        }

    }

    public void WriteData()
    {
        string range = sheetID + "!" + writeDataInRange;

        var valueRange = new ValueRange();
        var cellData = new List<object>();
        var arrows = new List<IList<object>>();
        foreach (var row in WriteDataFromUnity.rows)
        {
            cellData = new List<object>();
            foreach(var data in row.cellData)
            {
                cellData.Add(data);
            }
            arrows.Add(cellData);
        }

        valueRange.Values = arrows;

        var request = googleSheetService.Spreadsheets.Values.Append(valueRange, spreadSheetID, range);
        request.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.USERENTERED;
        var reponse = request.Execute();
    }

    public void DeleteData()
    {
        string range = sheetID + "!" + deleteDataInRange;

        var deleteData = googleSheetService.Spreadsheets.Values.Clear(new ClearValuesRequest(), spreadSheetID, range);
        deleteData.Execute();
    }

}
