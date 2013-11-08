#region Header
//
// (C) Copyright 2011-2013 by Autodesk, Inc.
//
// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE. AUTODESK, INC.
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is subject to
// restrictions set forth in FAR 52.227-19 (Commercial Computer
// Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii)
// (Rights in Technical Data and Computer Software), as applicable.
//
#endregion // Header

#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Application = Autodesk.Revit.ApplicationServices.Application;
using RvtOperationCanceledException = Autodesk.Revit.Exceptions.OperationCanceledException;
#endregion

namespace AdnPlugin.Revit.EstoreFile
{
  /// <summary>
  /// External Revit command to store the data of an
  /// arbitrary selected external file into Revit 
  /// extensible storage on a selected element.
  /// The data stored includes filename, original 
  /// folder name, and the file data itself.
  /// </summary>
  [Transaction( TransactionMode.Manual )]
  public class CmdStore : IExternalCommand
  {
    const string _schema_guid_string 
      = "7f2368b9-a45e-48b7-b5d5-446809c9ac19";

    static Guid _schema_guid 
      = new Guid( _schema_guid_string );

    /// <summary>
    /// Remember the last directory the user navigated to
    /// </summary>
    static string _initial_directory = null;

    /// <summary>
    /// Return an existing schema or
    /// warn user if none is present.
    /// </summary>
    public static Schema GetExistingSchema()
    {
      Schema schema = Schema.Lookup( _schema_guid );

      if( null == schema )
      {
        Util.InfoMessage( "No EstoreFile data is "
          + "present in this Revit document." );
      }
      return schema;
    }

    /// <summary>
    /// Retrieve an existing EstoreFile schema or 
    /// create a new one if it does not exist yet.
    /// </summary>
    /// <returns>EstoreFile schema</returns>
    static Schema GetSchema()
    {
      Schema schema = Schema.Lookup( CmdStore._schema_guid );

      if( null == schema )
      {
        SchemaBuilder schemaBuilder
          = new SchemaBuilder( _schema_guid );

        schemaBuilder.SetSchemaName( "EstoreFile" );

        // Allow anyone to read and write the object

        schemaBuilder.SetReadAccessLevel(
          AccessLevel.Public );

        schemaBuilder.SetWriteAccessLevel(
          AccessLevel.Public );

        // Create fields

        FieldBuilder fieldBuilder = schemaBuilder
          .AddSimpleField( "Filename", typeof( string ) );

        fieldBuilder.SetDocumentation( "File name" );

        fieldBuilder = schemaBuilder.AddSimpleField( 
          "Folder", typeof( string ) );

        fieldBuilder.SetDocumentation( "Original file folder path" );

        fieldBuilder = schemaBuilder.AddArrayField( 
          "Data", typeof( byte ) );

        fieldBuilder.SetDocumentation( "Stored file data" );

        // Register the schema

        schema = schemaBuilder.Finish();
      }
      return schema;
    }

    /// <summary>
    /// Store the given file data on the selected element.
    /// </summary>
    static bool EstoreFile( string filename, Element e )
    {
      Document doc = e.Document;

      bool rc = doc.IsModifiable;

      Debug.Assert( rc, 
        "expected a modifiable document" );

      if( rc )
      {
        Schema schema = GetSchema();

        // Check whether file data was 
        // already stored on this element

        Entity entity = e.GetEntity( schema );

        if( (null != entity) && entity.IsValid() )
        {
          string fn = entity.Get<string>( 
            schema.GetField( "Filename" ) );

          rc = Util.Question( string.Format(
            "This element already contains file data "
            + "for '{0}' in its extensible storage. "
            + "Overwrite existing data?", fn ) );
        }

        if( rc )
        {
          // Prepare the data to store

          byte[] data = File.ReadAllBytes( filename );

          string folder = Path.GetDirectoryName( filename );

          filename = Path.GetFileName( filename );

          // Save inital directory for next selection

          _initial_directory = folder;

          // Create an entity (object) for this schema (class)

          entity = new Entity( schema );

          // Set the values for this entity

          entity.Set<string>( schema.GetField( "Filename" ), filename );
          entity.Set<string>( schema.GetField( "Folder" ), folder );
          entity.Set<IList<byte>>( schema.GetField( "Data" ), data );

          // Store the entity on the element

          e.SetEntity( entity );

          Util.InfoMessage( string.Format(
            "Stored '{0}' file data on {1}",
            filename, Util.ElementDescription( e ) ) );
        }
      }
      return rc;
    }

    public Result Execute2(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements,
      bool storeOnType )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = Util.GetActiveDocument( uidoc, true );

      if( null == doc )
      {
        return Result.Failed;
      }

      Selection sel = uidoc.Selection;

      int n = sel.Elements.Size;

      if( 1 < n )
      {
        Util.InfoMessage( string.Format( 
          "{0} element{1} selected. Please select one "
          + "single element to store file data on.",
          n, Util.PluralSuffix( n ) ) );

        return Result.Failed;
      }

      OpenFileDialog dlg = new OpenFileDialog();
      
      dlg.Title = "Store File Data in Revit Extensible Storage";
      dlg.CheckFileExists = true;
      
      if( null != _initial_directory )
      {
        dlg.InitialDirectory = _initial_directory;
      }

      if( DialogResult.OK != dlg.ShowDialog() )
      {
        return Result.Cancelled;
      }

      Element e = null;

      if( 0 < n )
      {
        Debug.Assert( 1 == n,
          "we already checked for 1 < n above" );

        foreach( Element e2 in sel.Elements )
        {
          e = e2;
        }
      }
      else
      {
        try
        {
          Reference r = sel.PickObject(
            ObjectType.Element,
            "Please pick an element to store the file on: " );

          e = doc.GetElement( r.ElementId );
        }
        catch( RvtOperationCanceledException )
        {
          return Result.Cancelled;
        }
      }

      if( storeOnType )
      {
        ElementId id = e.GetTypeId();

        if( null == id
          || id.Equals( ElementId.InvalidElementId ) )
        {
          Util.InfoMessage( "The selected element "
            + "does not have an element type." );

          return Result.Failed;
        }

        e = doc.GetElement( id );
      }

      Transaction t = new Transaction( doc );

      t.Start( "Estore File" );

      EstoreFile( dlg.FileName, e );

      t.Commit();

      return Result.Succeeded;
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      return Execute2( commandData,
        ref message, elements, false );
    }
  }

  /// <summary>
  /// External Revit command to store the data of an
  /// arbitrary selected external file into Revit 
  /// extensible storage on the element type of a 
  /// selected element.
  /// The data stored includes filename, original 
  /// folder name, and the file data itself.
  /// </summary>
  [Transaction( TransactionMode.Manual )]
  public class CmdStoreOnType : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      return (new CmdStore()).Execute2( commandData, 
        ref message, elements, true );
    }
  }

  /// <summary>
  /// External Revit command to retrieve stored file 
  /// data from a selected element and recreate the 
  /// external file from it.
  /// </summary>
  [Transaction( TransactionMode.ReadOnly )]
  public class CmdRestore : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = Util.GetActiveDocument( uidoc, false );

      if( null == doc )
      {
        return Result.Failed;
      }

      Selection sel = uidoc.Selection;

      int n = sel.Elements.Size;

      if( 1 < n )
      {
        Util.InfoMessage( string.Format(
          "{0} element{1} selected. Please select one "
          + "single element to restore file data from.",
          n, Util.PluralSuffix( n ) ) );

        return Result.Failed;
      }

      Schema schema = CmdStore.GetExistingSchema();

      if( null == schema )
      {
        return Result.Failed;
      }

      Element e = null;

      if( 0 < n )
      {
        Debug.Assert( 1 == n,
          "we already checked for 1 < n above" );

        foreach( Element e2 in sel.Elements )
        {
          e = e2;
        }
      }
      else
      {
        try
        {
          Reference r = sel.PickObject(
            ObjectType.Element,
            "Please pick an element to restore the file from: " );

          e = doc.GetElement( r.ElementId );
        }
        catch( RvtOperationCanceledException )
        {
          return Result.Cancelled;
        }
      }

      Entity ent = e.GetEntity( schema );

      if( null == ent || !ent.IsValid() )
      {
        Util.InfoMessage( "No EstoreFile data "
          + "found on selected element." );

        return Result.Failed;
      }

      string filename = ent.Get<string>( schema.GetField( "Filename" ) );
      string folder = ent.Get<string>( schema.GetField( "Folder" ) );
      byte[] data = ent.Get<IList<byte>>( schema.GetField( "Data" ) ).ToArray<byte>();

      SaveFileDialog dlg = new SaveFileDialog();

      dlg.Title = "Restore File from Revit Extensible Storage";
      dlg.AddExtension = false;
      dlg.FileName = filename;

      if( Directory.Exists( folder ) )
      {
        dlg.InitialDirectory = folder;
      }

      if( DialogResult.OK != dlg.ShowDialog() )
      {
        return Result.Cancelled;
      }

      File.WriteAllBytes( dlg.FileName, data );

      return Result.Succeeded;
    }
  }

  /// <summary>
  /// External Revit command to list all 
  /// stored file data in the current document.
  /// Todo: support navigation to each element 
  /// containing EstoreFile data.
  /// </summary>
  [Transaction( TransactionMode.ReadOnly )]
  public class CmdList : IExternalCommand
  {
    /// <summary>
    /// A case independent string comparer class.
    /// This is required, or the map of filenames
    /// will consider C:\tmp\a.txt and C:\Tmp\a.txt 
    /// to be two different files.
    /// </summary>
    class PathComparer : IEqualityComparer<string>
    {
      public bool Equals( string x, string y )
      {
        return x.Equals( y, StringComparison
          .InvariantCultureIgnoreCase );
      }

      public int GetHashCode( string obj )
      {
        return obj.ToLower().GetHashCode();
      }
    }

    /// <summary>
    /// Return *all* elements in the document.
    /// </summary>
    static FilteredElementCollector GetAllElements(
      Document doc )
    {
      return new FilteredElementCollector( doc )
        .WherePasses( new LogicalOrFilter(
          new ElementIsElementTypeFilter( false ),
          new ElementIsElementTypeFilter( true ) ) );
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = Util.GetActiveDocument( uidoc, false );

      if( null == doc )
      {
        return Result.Failed;
      }

      Schema schema = CmdStore.GetExistingSchema();

      if( null == schema )
      {
        return Result.Succeeded;
      }

      FilteredElementCollector a = GetAllElements( doc );

      // Map stored file data to list of descriptions
      // of the BIM elements that it is stored on:

      Dictionary<string, List<string>> map
        = new Dictionary<string, List<string>>( 
          new PathComparer() );

      foreach( Element e in a )
      {
        Entity ent = e.GetEntity( schema );

        if( ent.IsValid() )
        {
          string filename = ent.Get<string>( schema.GetField( "Filename" ) );
          string folder = ent.Get<string>( schema.GetField( "Folder" ) );
          string path = Path.Combine( folder, filename );
          if( !map.ContainsKey( path ) )
          {
            map.Add( path, new List<string>( 1 ) );
          }
          map[path].Add( Util.ElementDescription( e ) );
        }
      }

      if( 0 == map.Count )
      {
        Util.InfoMessage( "No EstoreFile data is "
          + "present in this Revit document." );

        return Result.Succeeded;
      }

      List<string> keys = new List<string>( map.Keys );
      keys.Sort();

      string content = string.Empty;
      int nElements = 0;
      string elist;

      foreach( string path in keys )
      {
        nElements += map[path].Count;

        //ids = string.Join( ", ",
        //  map[path].ConvertAll<string>( id => 
        //    id.IntegerValue.ToString() ).ToArray() );

        //ids = string.Join( ", ", 
        //  map[path].ToArray() );

        elist = string.Join( "\r\n  ", 
          map[path].ToArray() );

        content += string.Format(
          "\r\n{0}: \r\n  {1}", path, elist );
      }

      int n = keys.Count;

      string instruction = string.Format(
        "{0} file{1} stored on {2} element{3}:",
        n, Util.PluralSuffix( n), nElements,
        Util.PluralSuffix( nElements ) );

      Util.InfoMessage( instruction, content );

      return Result.Succeeded;
    }
  }

  /// <summary>
  /// External Revit command to remove the file data 
  /// stored in extensible storage on selected elements.
  /// </summary>
  [Transaction( TransactionMode.Manual )]
  public class CmdRemove : IExternalCommand
  {
    class EstoreFileFilter : ISelectionFilter
    {
      Schema _schema;

      public EstoreFileFilter( Schema schema )
      {
        _schema = schema;
      }

      public bool AllowElement( Element e )
      {
        Entity ent = e.GetEntity( _schema );

        return null != ent && ent.IsValid();
      }

      public bool AllowReference( Reference r, XYZ p )
      {
        return true;
      }
    }

    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Document doc = Util.GetActiveDocument( uidoc, true );

      if( null == doc )
      {
        return Result.Failed;
      }

      Schema schema = CmdStore.GetExistingSchema();

      if( null == schema )
      {
        return Result.Succeeded;
      }

      Selection sel = uidoc.Selection;

      IList<Reference> refs;

      int n = sel.Elements.Size;

      if( 0 < n )
      {
        refs = new List<Reference>( n );

        foreach( Element e in sel.Elements )
        {
          refs.Add( new Reference( e ) );
        }
      }
      else
      {
        try
        {
          refs = sel.PickObjects(
            ObjectType.Element,
            new EstoreFileFilter( schema ),
            "Please select elements from which to "
            + "remove stored file data: " );
        }
        catch( RvtOperationCanceledException )
        {
          return Result.Cancelled;
        }
      }

      n = refs.Count;

      if( 0 == n )
      {
        Util.InfoMessage( "No elements selected." );

        return Result.Succeeded;
      }

      string q = string.Format(
        "Remove the file data stored in extensible "
        + "storage on {0} selected element{1}?",
        n, Util.PluralSuffix( n ) );

      if( Util.Question( q ) )
      {
        Transaction t = new Transaction( doc );

        t.Start( "Remove Stored Files" );

        Element e = null;

        foreach( Reference r in refs )
        {
          e = doc.GetElement( r.ElementId );

          e.DeleteEntity( schema );
        }

        t.Commit();
      }
      return Result.Succeeded;
    }
  }

  /// <summary>
  /// External Revit command to display the help file.
  /// </summary>
  [Transaction( TransactionMode.ReadOnly )]
  public class CmdHelp : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements )
    {
      Process.Start( App.HelpPath );
      return Result.Succeeded;
    }
  }
}

// C:\Program Files\Autodesk\Revit Architecture 2012\Program\Revit.exe
// C:\tmp\walls.rvt
// C:\Program Files\Autodesk\Revit Quasar Beta 2\Program\Revit.exe
// C:\tmp\estorewalls.rvt
