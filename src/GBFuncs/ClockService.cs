//==============================================
//              == GameBase ==
//    The Universal Retro-Gaming Frontend
//           and Database Utility
//----------------------------------------------
// Copyright (C) 1997-2010 James Burrows
// Email: gamebase.frontend@leesti.com
// Homepage: http://www.bu22.com/
//----------------------------------------------
//   GBFuncs:  General Functions used by GB
//----------------------------------------------
// This program is free software; you can
// redistribute it and/or modify it under the
// terms of the GNU General Public License as
// published by the Free Software Foundation;
// either version 2 of the License, or (at your
// option) any later version.
//
// This program is distributed in the hope that
// it will be useful, but WITHOUT ANY WARRANTY;
// without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR
// PURPOSE. See the GNU General Public License
// for more details.
//
// You should have received a copy of the GNU
// General Public License along with this
// program; if not, write to the Free Software
// Foundation, Inc., 59 Temple Place, Suite 330,
// Boston, MA 02111-1307 USA
//==============================================

using System.Diagnostics;

namespace GBFuncs
{
  /// <summary>
  /// Provides timing functionality for measuring elapsed time.
  /// Converted from VB6 clsClock class.
  /// </summary>
  public class ClockService
  {
    private readonly Stopwatch _stopwatch;

    public ClockService()
    {
      _stopwatch = new Stopwatch();
    }

    /// <summary>
    /// Starts or restarts the clock timer.
    /// </summary>
    public void StartClock()
    {
      _stopwatch.Restart();
    }

    /// <summary>
    /// Gets the time taken since StartClock was called, formatted as "0.00" seconds.
    /// </summary>
    public string TimeTaken
    {
      get
      {
        double seconds = _stopwatch.ElapsedMilliseconds / 1000.0;
        return seconds.ToString("0.00");
      }
    }

    /// <summary>
    /// Gets the time taken in milliseconds since StartClock was called.
    /// </summary>
    public long TimeTakenMilliseconds => _stopwatch.ElapsedMilliseconds;

    /// <summary>
    /// Stops the clock timer.
    /// </summary>
    public void StopClock()
    {
      _stopwatch.Stop();
    }

    /// <summary>
    /// Resets the clock timer to zero.
    /// </summary>
    public void ResetClock()
    {
      _stopwatch.Reset();
    }
  }
}
