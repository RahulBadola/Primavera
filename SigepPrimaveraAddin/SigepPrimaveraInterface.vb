Imports System.Runtime.InteropServices

<ComVisible(True), _
InterfaceType(ComInterfaceType.InterfaceIsIDispatch)> _
Public Interface SigepPrimaveraInterface

    <DispId(&H60020001)> _
      Sub Run(ByVal projId As String, ByVal direction As String, ByVal user As String)

End Interface
