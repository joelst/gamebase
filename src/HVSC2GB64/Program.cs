// GameBase - The Universal Retro-Gaming Frontend
// Copyright (C) 1997-2004 James Burrows - Migrated to .NET 2026
// HVSC2GB64: WinForms wizard application entry point

namespace HVSC2GB64;

static class Program
{
  /// <summary>
  /// The main entry point for the application.
  /// </summary>
  [STAThread]
  static void Main()
  {
    Application.EnableVisualStyles();
    Application.SetCompatibleTextRenderingDefault(false);
    Application.Run(new WizardForm());
  }
}
