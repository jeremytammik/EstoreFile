#region Copyright
// (C) Copyright 2011-2012 by Autodesk, Inc. 
//
// Permission to use, copy, modify, and distribute this software
// in object code form for any purpose and without fee is hereby
// granted, provided that the above copyright notice appears in
// all copies and that both that copyright notice and the limited
// warranty and restricted rights notice below appear in all
// supporting documentation.
//
// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS. 
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK,
// INC. DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL
// BE UNINTERRUPTED OR ERROR FREE.
//
// Use, duplication, or disclosure by the U.S. Government is
// subject to restrictions set forth in FAR 52.227-19 (Commercial
// Computer Software - Restricted Rights) and DFAR 252.227-7013(c)
// (1)(ii)(Rights in Technical Data and Computer Software), as
// applicable.
#endregion // Copyright

#region Namespaces
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Media.Imaging;
using System.Collections.Generic;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Diagnostics;
#endregion

namespace AdnPlugin.Revit.EstoreFile
{
  class App : IExternalApplication
  {
    const string _help_filename = "EstoreFile.htm";
    static string _help_path = null;

    const string _name = Util.Caption;

    static string _text = _name; // AssemblyProduct

    static string _namespace_prefix
      = typeof( App ).Namespace + ".";

    static string[] _commands = new string[] { 
      "Store", "StoreOnType", "Restore", 
      "List", "Remove", "Help" };

    #region Tooltip string constants
    static string[] _tips = new string[] { 
      "Store the data of an external file into Revit extensible storage on a selected element.", 
      "Store the data of an external file into Revit extensible storage on the element type of a selected element.", 
      "Restore external file from extensible storage on a selected element.",
      "List file data stored in extensible storage.", 
      "Remove file data stored in extensible storage.",
      "Display the help file.",
    };

    static string[] _long_tips = new string[] { 
      "This command prompts you to select an external file and a Revit BIM element and stores the file data into Revit extensible storage.", 
      "This command prompts you to select an external file and a Revit BIM element and stores the file data into the Revit extensible storage on the selected element's element type.", 
      "This command prompts you to select an element and retrieves the stored external file from its extensible storage.",
      "This command lists all the file file data stored in extensible storage on all elements in the current active document.", 
      "This command prompts you to select one or more BIM elements and removes the file data stored in their extensible storage.",
      "This command displays the help file. It is only available in Revit 2012. In Revit 2013, you can hover over a command button and press the F1 key instead.",
    };

    static string _long_description_tooltip_overview
      = "\r\n\r\n\r\n\r\nAdd-in Overview\r\n\r\n"
      + "This add-in stores and manages the contents of external files in Revit extensible storage on selected elements in the BIM model. "
      + "It provides the following commands:\r\n\r\n"
      + "Store: Store the data of an external file into Revit extensible storage on a selected element.\r\n\r\n"
      + "StoreOnType: Store the data of an external file into Revit extensible storage on the element type of a selected element.\r\n\r\n"
      + "Restore: Restore external file from extensible storage on a selected element.\r\n\r\n"
      + "List: List file data stored in extensible storage.\r\n\r\n"
      + "Remove: Remove file data stored in extensible storage.\r\n\r\n"
      + "Help: Display the help file.\r\n\r\n";
    #endregion // Tooltip string constants

    #region Assembly attribute accessors
    /// <summary>
    /// Shortcut to retrieve executing assembly
    /// </summary>
    static public Assembly ExecutingAssembly
    {
      get
      {
        return Assembly.GetExecutingAssembly();
      }
    }

    static object GetFirstCustomAttribute( Type t )
    {
      Assembly a = ExecutingAssembly;

      object[] attributes
        = a.GetCustomAttributes( t, false );

      return ( 0 < attributes.Length )
        ? attributes[0]
        : null;
    }

    static public string AssemblyProduct
    {
      get
      {
        object a = GetFirstCustomAttribute( 
          typeof( AssemblyProductAttribute ) );

        return ( null == a )
          ? string.Empty
          : ( (AssemblyProductAttribute) a ).Product;
      }
    }

    static string AssemblyDescription
    {
      get
      {
        object a = GetFirstCustomAttribute( 
          typeof( AssemblyDescriptionAttribute ) );

        return ( null == a )
          ? string.Empty
          : ( (AssemblyDescriptionAttribute) a ).Description;
      }
    }
    #endregion // Assembly attribute accessors

    /// <summary>
    /// Load a new icon bitmap from embedded resources.
    /// For the BitmapImage, make sure you reference 
    /// WindowsBase and PresentationCore, and import 
    /// the System.Windows.Media.Imaging namespace. 
    /// </summary>
    BitmapImage NewBitmapImage(
      Assembly a,
      string imageName )
    {
      // To read from an external file:
      //return new BitmapImage( new Uri(
      //  Path.Combine( _imageFolder, imageName ) ) );

      Stream s = a.GetManifestResourceStream(
         _namespace_prefix + "Icon." + imageName );

      BitmapImage img = new BitmapImage();

      img.BeginInit();
      img.StreamSource = s;
      img.EndInit();

      return img;
    }

    /// <summary>
    /// Return the full path to the add-in help file
    /// </summary>
    public static string HelpPath
    {
      get { return _help_path; }
    }

    #region Revit 2013 contextual help support
    /// <summary>
    /// Retrieve the Revit 2013 ContextualHelp class 
    /// type if it exists.
    /// </summary>
    static Type _contextual_help_type = null;

    /// <summary>
    /// Call this method once only to set the 
    /// ContextualHelp class type if it exists.
    /// This can be used for a runtime detection 
    /// whether we are running in Revit 2012 or 2013.
    /// </summary>
    static void SetContextualHelpType()
    {
      // Create an instance of the RevitAPIUI.DLL 
      // assembly to query its type information:

      Assembly rvtApiUiAssembly 
        = Assembly.GetAssembly( typeof( UIDocument ) );
      
      _contextual_help_type = rvtApiUiAssembly.GetType( 
        "Autodesk.Revit.UI.ContextualHelp" );
    }

    /// <summary>
    /// Return true if the running version of the 
    /// Revit API supports the ContextualHelp class,
    /// i.e. are running inside Revit 2013 or later.
    /// </summary>
    static bool HasContextualHelp
    {
      get
      {
        return null != _contextual_help_type;
      }
    }

    /// <summary>
    /// Assign the add-in help file to the 
    /// given ribbon push button contextual help.
    /// This method checks for Revit 2013 at runtime
    /// and uses .NET Reflection to access the 
    /// ContextualHelp class and SetContextualHelp
    /// method which are not available in the 
    /// Revit 2012 API.
    /// </summary>
    void AddContextualHelp( PushButtonData d, int i )
    {
      // This is the normal code uses
      // when directly referencing the
      // Revit 2013 API assemblies:
      //
      //ContextualHelp ch = new ContextualHelp(
      //  ContextualHelpType.ChmFile, helppath );
      //
      //d.SetContextualHelp( ch );

      // Invoke constructor:

      ConstructorInfo[] ctors 
        = _contextual_help_type.GetConstructors();

      const int contextualHelpType_ChmFile = 3;

      object instance = ctors[0].Invoke(
        new object[] { 
        contextualHelpType_ChmFile, 
        HelpPath } ); // + "#" + i.ToString()

      // Set the help topic URL

      PropertyInfo property = _contextual_help_type
        .GetProperty( "HelpTopicUrl" );

      property.SetValue( instance, i.ToString(), null );

      // Invoke SetContextualHelp method:

      Type pbdType = d.GetType();

      MethodInfo method = pbdType.GetMethod( 
        "SetContextualHelp" );

      method.Invoke( d, new object[] { instance } );
    }
    #endregion // Revit 2013 contextual help support

    public Result OnStartup( UIControlledApplication a )
    {
      int i;

      SetContextualHelpType();

      if( HasContextualHelp )
      {
        // Remove the description of the help command 
        // from the long tooltip add-in overview

        i = _long_description_tooltip_overview
          .IndexOf( "\r\n\r\nHelp: " );

        Debug.Assert( 0 < i, "expected to find "
          + "description of help command in long "
          + "description tootip overview" );

        _long_description_tooltip_overview 
          = _long_description_tooltip_overview
            .Substring( 0, i );
      }

      ControlledApplication c = a.ControlledApplication;

      Debug.Print( string.Format( 
        "EstoreFile: Running in {0} {1}.{2}, "
        + "so contecxtual help is {3}available.", 
        c.VersionName, c.VersionNumber, c.VersionBuild, 
        (HasContextualHelp ? "" : "not ") ) );

      Assembly exe = Assembly.GetExecutingAssembly();
      string dllpath = exe.Location;
      string dir = Path.GetDirectoryName( dllpath );
      _help_path = Path.Combine( dir, _help_filename );

      if( !File.Exists( _help_path ) )
      {
        string s = "Please ensure that the EstoreFile "
          + "help file {0} is located in the same "
          + " folder as the add-in assembly DLL:\r\n"
          + "'{1}'.";

        System.Windows.MessageBox.Show( 
          string.Format( s, _help_filename, dir ),
          _name, System.Windows.MessageBoxButton.OK, 
          System.Windows.MessageBoxImage.Error );

        return Result.Failed;
      }

      string className = GetType().FullName.Replace(
        "App", "Cmd" );

      RibbonPanel rp = a.CreateRibbonPanel( _text );

      i = 0;

      foreach( string cmd in _commands )
      {
        if( HasContextualHelp && cmd.Equals( "Help" ) )
        {
          break;
        }

        PushButtonData d = new PushButtonData(
          cmd, cmd, dllpath, className + cmd );

        d.ToolTip = _tips[i];

        d.Image = NewBitmapImage( exe, 
          cmd + "16.png" );

        d.LargeImage = NewBitmapImage( exe, 
          cmd + "32.png" );

        d.LongDescription = _long_tips[i] 
          + _long_description_tooltip_overview;

        if( HasContextualHelp )
        {
          AddContextualHelp( d, ++i );
        }

        rp.AddItem( d );
      }
      return Result.Succeeded;
    }

    public Result OnShutdown( UIControlledApplication a )
    {
      return Result.Succeeded;
    }
  }
}
