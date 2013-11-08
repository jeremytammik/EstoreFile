#region Header
//
// (C) Copyright 2011 by Autodesk, Inc.
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
using System.Diagnostics;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
#endregion

namespace AdnPlugin.Revit.EstoreFile
{
  class Util
  {
    public const string Caption = "Extensible Storage";

    /// <summary>
    /// Return an English plural suffix for the given 
    /// number of items, i.e. 's' for zero or more 
    /// than one, and nothing for exactly one.
    /// </summary>
    public static string PluralSuffix( int n )
    {
      return 1 == n ? "" : "s";
    }

    /// <summary>
    /// Return a full stop dot for zero items,
    /// and a colon for more than zero.
    /// </summary>
    public static string DotOrColon( int n )
    {
      return 0 == n ? "." : ":";
    }

    /// <summary>
    /// Display an informational message 
    /// in a Revit task dialogue.
    /// </summary>
    public static void InfoMessage( string instruction )
    {
      Debug.Print( instruction );

      TaskDialog a = new TaskDialog( Caption );

      a.MainInstruction = instruction;
      a.Show();
    }

    /// <summary>
    /// Display an informational message 
    /// in a Revit task dialogue.
    /// </summary>
    public static void InfoMessage(
      string instruction,
      string content )
    {
      Debug.Print( instruction );
      Debug.Print( content );

      TaskDialog a = new TaskDialog( Caption );

      a.MainInstruction = instruction;
      a.MainContent = content;
      a.Show();
    }

    /// <summary>
    /// Ask a question and return a Boolean answer.
    /// </summary>
    public static bool Question( string question )
    {
      Debug.Print( question );

      TaskDialog a = new TaskDialog( Caption );
      a.MainInstruction = question;
      a.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
      a.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No;
      a.DefaultButton = TaskDialogResult.No;

      bool rc = TaskDialogResult.Yes == a.Show();

      Debug.Print( rc ? "Yes" : "No" );

      return rc;
    }

    /// <summary>
    /// Return active document or 
    /// warn user if there is none.
    /// Optionally, require the document to be 
    /// modifiable.
    /// </summary>
    public static Document GetActiveDocument( 
      UIDocument uidoc,
      bool requireModifiable )
    {
      Document doc = uidoc.Document;

      if( null == doc )
      {
        Util.InfoMessage( "Please run this command "
          + "in an activate Revit document." );
      }
      else if( requireModifiable && doc.IsReadOnly )
      {
        Util.InfoMessage( "The active document is "
          + "read-only. This command requires a "
          + "modifiable Revit document." );

        doc = null;
      }
      return doc;
    }

    /// <summary>
    /// Return a string describing the given element:
    /// .NET type name,
    /// category name,
    /// family and symbol name for a family instance,
    /// element id and element name.
    /// </summary>
    public static string ElementDescription( Element e )
    {
      if( null == e )
      {
        return "<null>";
      }

      // For a wall, the element name equals the
      // wall type name, which is equivalent to the
      // family name ...
      
      FamilyInstance fi = e as FamilyInstance;

      string typeName = e.GetType().Name;

      string categoryName = ( null == e.Category )
        ? string.Empty
        : e.Category.Name + " ";

      string familyName = ( null == fi )
        ? string.Empty
        : fi.Symbol.Family.Name + " ";

      string symbolName = ( null == fi 
        || e.Name.Equals( fi.Symbol.Name ) )
          ? string.Empty
          : fi.Symbol.Name + " ";

      return string.Format( "{0} {1}{2}{3}<{4} {5}>",
        typeName, categoryName, familyName, symbolName, 
        e.Id.IntegerValue, e.Name );
    }
  }
}
