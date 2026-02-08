//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1999-2010 James Burrows
// Migrated to .NET 2026
//----------------------------------------------
//   About Dialog
//==============================================

using System.Reflection;

namespace GameBase;

public partial class AboutForm : Form
{
  public AboutForm()
  {
    InitializeComponent();
    LoadVersionInfo();
  }

  private void LoadVersionInfo()
  {
    var assembly = Assembly.GetExecutingAssembly();
    var version = assembly.GetName().Version;
    var informationalVersion = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;

    lblVersion.Text = $"Version {informationalVersion ?? version?.ToString() ?? "1.0.0"}";
    lblCopyright.Text = "Copyright © 1999-2010 James Burrows\nMigrated to .NET 2026";

    txtAbout.Text = @"GameBase - The Universal Retro-Gaming Frontend and Database Utility

GameBase is a comprehensive database and launcher for retro games, allowing you to organize, browse, and play your vintage game collection.

Features:
• Comprehensive game database with metadata
• Screenshot and extras management
• Advanced filtering and custom views
• GEMUS script engine integration
• Cross-platform support (.NET 8)

Technology Stack:
• .NET 8 Windows Forms
• SQLite database
• GEMUS script engine
• 7-Zip / SharpCompress archive support

Originally created by James Burrows
Modernized and migrated to .NET in 2026

This software is provided as-is under the original GameBase license.";
  }

  private void btnOk_Click(object? sender, EventArgs e)
  {
    Close();
  }
}
