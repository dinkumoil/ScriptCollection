'==============================================================================
'
' Class to store Notepad++ plugin data
'
' Author: Andreas Heim
' Date:   29.09.2019
'
' This program is free software; you can redistribute it and/or modify it
' under the terms of the GNU General Public License version 3 as published
' by the Free Software Foundation.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along
' with this program; if not, write to the Free Software Foundation, Inc.,
' 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'==============================================================================

Class clsNppPlugin
  Private strDisplayName
  Private strFolderName
  Private strRepository
  Private strVersion
  Private strInstalledVersion
  Private strInstallPath
  Private bolIsInstalled
  Private bolIsUpdateAvailable

  '----------------------------------------------------------------------------
  'Constructor
  '----------------------------------------------------------------------------
  Private Sub Class_Initialize()
    strDisplayName      = ""
    strFolderName       = ""
    strRepository       = ""
    strVersion          = ""
    trInstalledVersion  = ""
    trInstallPath       = ""
    olIsInstalled       = ""
    olIsUpdateAvailable = ""
  End Sub

  '----------------------------------------------------------------------------
  'Destructor
  '----------------------------------------------------------------------------
  Private Sub Class_Terminate()
  End Sub

  '----------------------------------------------------------------------------
  'Getter/Setter
  '----------------------------------------------------------------------------
  Public Property Let DisplayName(strValue)
    strDisplayName = strValue
  End Property

  Public Property Get DisplayName
    DisplayName = strDisplayName
  End Property

  Public Property Let FolderName(strValue)
    strFolderName = strValue
  End Property

  Public Property Get FolderName
    FolderName = strFolderName
  End Property

  Public Property Let Repository(strValue)
    strRepository = strValue
  End Property

  Public Property Get Repository
    Repository = strRepository
  End Property

  Public Property Let Version(strValue)
    strVersion = strValue
  End Property

  Public Property Get Version
    Version = strVersion
  End Property

  Public Property Let InstalledVersion(strValue)
    strInstalledVersion = strValue
  End Property

  Public Property Get InstalledVersion
    InstalledVersion = strInstalledVersion
  End Property

  Public Property Let InstallPath(strValue)
    strInstallPath = strValue
  End Property

  Public Property Get InstallPath
    InstallPath = strInstallPath
  End Property

  Public Property Let IsInstalled(bolValue)
    bolIsInstalled = bolValue
  End Property

  Public Property Get IsInstalled
    IsInstalled = bolIsInstalled
  End Property

  Public Property Let IsUpdateAvailable(bolValue)
    bolIsUpdateAvailable = bolValue
  End Property

  Public Property Get IsUpdateAvailable
    IsUpdateAvailable = bolIsUpdateAvailable
  End Property
End Class
