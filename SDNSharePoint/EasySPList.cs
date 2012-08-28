using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml;
using System.Net;

namespace SDNSharePoint
{
    public class EasySPList : IDisposable
    {
      private _SPList.Lists MyList;
      
      private DataSet _DS = null;
      private bool _isDisposed = false;
      private string NeedFields = null;

      private XmlDocument doc = null;
      private XmlElement FieldsElement = null;


      public bool CleanRetrunValues = false;



      public string URL
      {
          set
          {
              MyList.Url = value;
          }
      }

      private void SetConstructor(string URL, bool CleanRetrunValues, NetworkCredential CustomCredentials)
      {
          if (!URL.StartsWith("http"))
              URL = "http://" + URL;

          MyList = new _SPList.Lists();
          MyList.Url = URL + "/_vti_bin/Lists.asmx";
          MyList.Credentials = CustomCredentials;

          _DS = new DataSet();

          doc = new XmlDocument();

          FieldsElement = doc.CreateElement("ViewFields");

      }


      public EasySPList(string URL)
      {
          SetConstructor(URL, CleanRetrunValues, GetCredentials());
      }

      public EasySPList(string URL, bool CleanRetrunValues)
      {
          SetConstructor(URL, CleanRetrunValues, GetCredentials());
      }

      public EasySPList(string URL, bool CleanRetrunValues, NetworkCredential CustomCredentials)
      {
          SetConstructor(URL, CleanRetrunValues, CustomCredentials);
      }

      private NetworkCredential GetCredentials()
      {

          return CredentialCache.DefaultNetworkCredentials;

      }


      #region Query

      /// <summary>
      /// Returns all rows in the site.
      /// </summary>
      /// <param name="ListName">The list name you want the data from</param>
      /// <returns>Returns a dataset the rows</returns>
      public DataSet SearchList(string ListName)
      {
          return SearchList(ListName, null, null, null);
      }

      /// <summary>
      /// Returns all rows in the site or the max number of rows you say you want back if number of rows are more then you request
      /// </summary>
      /// <param name="ListName">The list name you want the data from.</param>
      /// <param name="RowLimit">The max number of rows you want back.</param>
      /// <returns>Returns a dataset the rows</returns>
      public DataSet SearchList(string ListName, string RowLimit)
      {
          return SearchList(ListName, RowLimit, null, null);
      }

      /// <summary>
      /// Returns all rows in the site or the max number of rows you say you want back if number of rows are more then you request
      /// </summary>
      /// <param name="ListName">The list name you want the data from.</param>
      /// <param name="RowLimit">The max number of rows you want back.</param>
      /// <param name="Fields">The fields you want to be returned in your query.</param>
      /// <returns>Returns a dataset the rows</returns>
      public DataSet SearchList(string ListName, string RowLimit, XmlElement Fields)
      {
          return SearchList(ListName, RowLimit, FieldsElement, null);
      }

      /// <summary>
      /// Returns all rows in the site or the max number of rows you say you want back if number of rows are more then you request. 
      /// Also, you can pass in a list of fields with a specil charter and the function will create the CAML field code for you.
      /// </summary>
      /// <param name="ListName">The list name you want the data from.</param>
      /// <param name="RowLimit">The max number of rows you want back.</param>
      /// <param name="Fields">The list of fields you want to be returned in your query. e.x: "Field1;Field2;Field3"</param>
      /// <param name="SplitChar">The specail char you want to split of the fields by.  e.x: ';'</param>
      /// <returns>Returns a dataset the rows</returns>
      public DataSet SearchList(string ListName, string RowLimit, string Fields, char SplitChar)
      {
          CreatCAMLFeilds(Fields, SplitChar);

          return SearchList(ListName, RowLimit, FieldsElement, null);
      }


      /// <summary>
      /// Returns all rows in the site or the max number of rows you say you want back if number of rows are more then you request. 
      /// </summary>
      /// <param name="ListName">The list name you want the data from.</param>
      /// <param name="RowLimit">The max number of rows you want back.</param>
      /// <param name="Fields">The list of fields you want to be returned in your query. e.x: "Field1;Field2;Field3"</param>
      /// <param name="SplitChar">The specail char you want to split of the fields by.  e.x: ';'</param>
      /// <param name="Query">CAML code  of your query/sql statement</param>
      /// <returns>Returns a dataset the rows</returns>
      public DataSet SearchList(string ListName, string RowLimit, string Fields, string Query, char SplitChar)
      {

          CreatCAMLFeilds(Fields, SplitChar);

          return SearchList(ListName, RowLimit, FieldsElement, Query);
      }

      /// <summary>
      /// Returns all rows in the site or the max number of rows you say you want back if number of rows are more then you request. 
      /// </summary>
      /// <param name="ListName">Name of the list you want.</param>
      /// <param name="RowLimit">Number of rows that will be returned.</param>
      /// <param name="Fields">What fields you want returned.  If null it will return only fields with a value.  List default items will not be returned unless you pass them in.</param>
      /// <param name="Query">The query/search you want returned.</param>
      /// <returns>Returns a dataset of the return values.   </returns>
      public DataSet SearchList(string ListName,
                              string RowLimit,
                              XmlElement Fields,
                              string Query)
      {
          XmlNode node = null;

          //XmlElement FieldsElement = null;
          XmlElement batchElement = null;

          XmlTextReader reader = null;
          _DS = new DataSet();
         

          try
          {

              ///* Specify methods for the batch post using CAML. In each method  include the ID of the item to update and the value to place in the specified column.*/
              //if (!string.IsNullOrEmpty(Fields))
              //{
              //    FieldsElement = doc.CreateElement("ViewFields");
              //    FieldsElement.InnerXml = Fields;
              //}

              if (!string.IsNullOrEmpty(Query))
              {
                  batchElement = doc.CreateElement("Query");
                  batchElement.InnerXml = Query;
              }


              // Retrieve raw XML from SharePoint web service

              node = MyList.GetListItems(ListName, null, batchElement, Fields, RowLimit, null, null);

              if (!(node.OuterXml.ToLower().IndexOf("itemcount=\"0\"") >= 0))
              {
                  string tempOuterXML = string.Empty;

                  if (CleanRetrunValues)
                  {
                      tempOuterXML = node.OuterXml.Replace("ows_", "");
                      tempOuterXML = tempOuterXML.Replace("row", ListName.Replace(" ", ""));
                      tempOuterXML = tempOuterXML.Replace("_x0020_", "");
                      tempOuterXML = tempOuterXML.Replace("1;#", "");

                  }
                  else
                      tempOuterXML = node.OuterXml;

                  // Add raw XML into XmlTextReader
                  reader = new XmlTextReader(tempOuterXML, XmlNodeType.Element, null);

                  // Load XmlTextReader into DataSet
                  _DS.ReadXml(reader);

                  if (CleanRetrunValues)
                      if (_DS.Tables.Count.Equals(2))
                          if (_DS.Tables[1].TableName.Equals("row"))
                              _DS.Tables[1].TableName = ListName;



              }

          }
          catch (Exception e)
          {
              throw e;
          }
          finally
          {
              #region CleanUp

              if (reader != null)
                  reader.Close();


              reader = null;
              node = null;
              //doc = null;

              #endregion
          }

          return _DS;

      }//GetListDataSet



      #endregion

      #region Create Rows

      /// <summary>
      /// Add's one row to a SharePoint list.  The program will create your CAML code for you.   You will need to split the Field and Value with a "=".  
      /// e.x: AddARowToList("testlist","field1=Value;field2=Value;field3=value",';');
      /// </summary>
      /// <param name="ListName">The list name you want to use.</param>
      /// <param name="FieldsAndValues">The fields and values you want to add to the list.  E.X: field1=Value;field2=Value;field3=Three</param>
      /// <param name="CharSplit">A special charter that will split the fields up.</param>
      public string AddARowToList(string ListName, string FieldsAndValues, Char CharSplit)
      {
          XmlNode node = null;
          XmlTextReader reader = null;
          XmlElement batchElement = null;

          //_DS.Clear();


          CreatCAMLForAdd(FieldsAndValues, CharSplit);

          //XmlDocument doc = new XmlDocument();
          XmlElement Batch = doc.CreateElement("Batch");

          Batch.SetAttribute("OnError", "Return");

          Batch.InnerXml += string.Format("<Method ID=\"1\" Cmd=\"New\">" +
                                          NeedFields +
                                          "</Method>");

          node = MyList.UpdateListItems(ListName, Batch);
          return node.OuterXml;

          if (!(node.OuterXml.ToLower().IndexOf("itemcount=\"0\"") >= 0))
          {

              string tempOuterXML = string.Empty;

              if (CleanRetrunValues)
              {
                  tempOuterXML = node.OuterXml;
                  tempOuterXML = tempOuterXML.Replace("ows_", "");
                  tempOuterXML = tempOuterXML.Replace("z:row", ListName);
                  tempOuterXML = tempOuterXML.Replace("_x0020_", "");
                  tempOuterXML = tempOuterXML.Replace("1;#", "");

              }
              else
                  tempOuterXML = node.OuterXml;

              return tempOuterXML;
              // Add raw XML into XmlTextReader
              reader = new XmlTextReader(tempOuterXML, XmlNodeType.Element, null);

              // Load XmlTextReader into DataSet
              //_DS.ReadXml(reader);

              if (CleanRetrunValues)
                  if (_DS.Tables.Count.Equals(2))
                      if (_DS.Tables[1].TableName.Equals("row"))
                          _DS.Tables[1].TableName = ListName;



          }

          return null;




      }

      #endregion

      #region Update

      /// <summary>
      /// Updates the row(s) that you need updated.
      /// </summary>
      /// <param name="ListName">The name of the list you want to update.</param>
      /// <param name="UpdateCAML">The Update CAML code.</param>
      public string UpdateEntry(string ListName, string UpdateCAML)
      {

          
          XmlTextReader reader = null;
          _DS = new DataSet();
          
          XmlElement Batch = doc.CreateElement("Batch");

          Batch.SetAttribute("OnError", "Return");

          Batch.InnerXml = UpdateCAML;

          XmlNode node = null;

          try
          {
              node = MyList.UpdateListItems(ListName, Batch);


          }
          catch (Exception err)
          {
              throw new Exception(err.InnerException.Message);
          }


          return node.OuterXml;

         


      }

      #endregion

      private void CreatCAMLFeilds(string Fields, char SplitChar)
      {
          FieldsElement = doc.CreateElement("ViewFields");
          NeedFields = string.Empty;

          if (string.IsNullOrEmpty(Fields))
              NeedFields = null;
          else
          {
              foreach (string field in Fields.Split(SplitChar))
              {
                  NeedFields += "<FieldRef Name='" + field.Replace(" ", "_x0020_") + "' />";
              }
              FieldsElement.InnerXml = NeedFields;
          }


      }

      private void CreatCAMLForAdd(string Fields, char SplitChar)
      {
          FieldsElement = doc.CreateElement("ViewFields");
          NeedFields = string.Empty;

          if (SplitChar == null && Fields.Length > 0)
              NeedFields = Fields;
          else if (SplitChar == null && !string.IsNullOrEmpty(Fields))
              FieldsElement = null;
          else
          {
              foreach (string field in Fields.Split(SplitChar))
              {
                  string[] temp = field.Split("=".ToCharArray(), 2);

                  NeedFields += "<Field Name=\"" + temp[0].Replace(" ", "_x0020_") + "\">" + temp[1] + "</Field>";
              }
              //FieldsElement.InnerXml = NeedFields;
          }


      }


      




      #region Dispose Area
      protected virtual void Dispose(bool disposing)
        {
            if (!_isDisposed) // only dispose once!
            {
                if (disposing)
                {
                    MyList.Dispose();
                    _DS.Clear();
                    _DS.Dispose();

                    MyList = null;
                    _DS = null;

                }
                // Code to dispose the un-managed resources of the class
            }
            _isDisposed = true;
        }

        public void Dispose()
        {
            Dispose(true);
            //If dispose is called already then say GC to skip finalize on this instance.
            GC.SuppressFinalize(this);
        }

        ~EasySPList()
        {
            Dispose(false);
        }

     #endregion




    }





}
