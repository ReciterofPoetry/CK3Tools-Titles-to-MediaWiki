Imports System.IO
Imports System.IO.Compression
Imports System.Runtime.CompilerServices
Imports System.Text
Friend Module Props
    Property BaseDir As String = Environment.CurrentDirectory
    Property GS As String = $"{BaseDir}\gamestate"
    Property GameDir As String
    Property Config As String = $"{Path.GetTempPath()}{Path.DirectorySeparatorChar}CK3Tools.txt"
    Property OutputName As String
    Property SaveDate As String
    Property Mods As New List(Of String)
    Property ModPaths As New List(Of String)
    Property ModNames As New List(Of String)
    Property ModsFolder As String
End Module
Module Program
    Const e As Byte = 0
    Const k As Byte = 1
    Const d As Byte = 2
    Const c As Byte = 3
    Const b As Byte = 4
#Disable Warning IDE0044 ' Add readonly modifier
    Dim GameConceptLocalisations As New Hashtable()
    Dim SavedLocalisation As New Hashtable()

    Dim Titles As New Dictionary(Of Integer, String) 'Keys matched with title ids.
    Dim Names As New Dictionary(Of Integer, String) 'Keys matched with names.
    Dim Colours As New Dictionary(Of Integer, String) 'Keys matched with RGB colours.
    Dim Capitals As New Dictionary(Of Integer, String) 'Keys matched with their capitals, no baronies.

    Dim DeJure, Titular As New Dictionary(Of Byte, List(Of Integer)) From {{e, New List(Of Integer)}, {k, New List(Of Integer)}, {d, New List(Of Integer)}} 'Each dictionary has sublists for each of empire, kingdom, and duchy titles. These will be used later as the primary references for the tables. Same for the following list.
    Dim Counties As New List(Of Integer) 'List of integer.

    Dim Vassals As New Dictionary(Of Byte, Dictionary(Of Integer, List(Of String))) From {{k, New Dictionary(Of Integer, List(Of String))}, {d, New Dictionary(Of Integer, List(Of String))}, {c, New Dictionary(Of Integer, List(Of String))}, {b, New Dictionary(Of Integer, List(Of String))}} 'Vassals sorted according to vassal title ranks, then the lieges. So Vassals(k) is the kingdom vassals, and then you get the kingdom vassals of any given empire from the dictionary that Vassals(k) references.
    Dim VassalCounts As New Dictionary(Of Byte, Dictionary(Of Integer, Integer)) From {{k, New Dictionary(Of Integer, Integer)}, {d, New Dictionary(Of Integer, Integer)}, {c, New Dictionary(Of Integer, Integer)}, {b, New Dictionary(Of Integer, Integer)}} 'Same as above except it stores the total number of those vassals that any given title has.
    Dim LargestCounty As New Dictionary(Of Integer, Integer)

    Dim Lieges As New Dictionary(Of Byte, Dictionary(Of Integer, String)) From {{e, New Dictionary(Of Integer, String)}, {k, New Dictionary(Of Integer, String)}, {d, New Dictionary(Of Integer, String)}} 'Similar to above except inversed to get lieges. So Lieges(e) is the empire lieges, then you get the get the empire liege of any given title from the dictionary that Lieges(e) references.


    'Dim SpecialBuildings As New Dictionary(Of Byte, Dictionary(Of Integer, List(Of String)))
    Dim Modifiers, SpecialBuildings As New Dictionary(Of Integer, Dictionary(Of String, Integer))

    Dim Faiths, Cultures As New Dictionary(Of Integer, String) 'Stores the faith and culture of counties.
    Dim CountyDev, TotalDev, HighestDev As New Dictionary(Of Integer, Integer) 'Stores the development of counties, the total dev of county vassals under a title, and the highest individual dev from the counties under a title, respectively.
#Enable Warning IDE0044 ' Add readonly modifier
    Sub Main()

        Dim StartTime As DateTime = Now

        SetUp()
        CollectLocalisations()

        Dim LandedTitlesFiles As New List(Of String) From {GameDir & "/common/landed_titles/00_landed_titles.txt"}
        For Each ModPath In ModPaths
            Dim RawFiles() As String
            Dim LandedTitlesPath As String = Path.Combine(ModPath, "common/landed_titles/")
            If Directory.Exists(LandedTitlesPath) Then
                RawFiles = Directory.GetFiles(LandedTitlesPath)
            End If
#Disable Warning BC42104 ' Variable is used before it has been assigned a value
            If Not RawFiles.Length = 0 Then
#Enable Warning BC42104 ' Variable is used before it has been assigned a value
                LandedTitlesFiles.AddRange(RawFiles)
            End If
        Next

        Dim LandedTitlesData As New Dictionary(Of String, List(Of String))
        Dim Provinces As New Dictionary(Of Integer, String)

        For Each TextFile In LandedTitlesFiles
            Using SR As New StreamReader(TextFile)
                Dim LineData As String
                Dim Title As String = ""
                Dim Data As New List(Of String)
                Dim Index As Integer
                While Not SR.EndOfStream
                    LineData = SR.ReadLine.TrimStart
                    With LineData
                        If .StartsWith("e_") OrElse .StartsWith("k_") OrElse .StartsWith("d_") OrElse .StartsWith("c_") OrElse .StartsWith("b_") Then
                            If Not Data.Count = 0 Then
                                LandedTitlesData(Title) = Data.ToList
                                Data.Clear()
                            End If
                            Index = .IndexOf("="c)
                            Title = LineData.Substring(0, Index).TrimEnd
                            If Not LandedTitlesData.ContainsKey(Title) Then
                                LandedTitlesData.Add(Title, New List(Of String))
                            End If
                        ElseIf .StartsWith("color") AndAlso Not .StartsWith("color2") AndAlso Not Title.StartsWith("b"c) AndAlso Not Data.Contains(LineData) Then
                            Data.Add(LineData)
                        ElseIf .Contains("cn_") AndAlso .Contains("="c) AndAlso .IndexOf("="c) < .IndexOf("cn") Then
                            Index = .IndexOf("="c) + 1
                            LineData = .Substring(Index).Trim
                            If Not Data.Contains(LineData) Then
                                Data.Add(LineData)
                            End If
                        ElseIf .StartsWith("province") AndAlso Not Data.Contains(LineData) Then
                            Dim Province As Integer = .Split("="c, 2).Last.Trim.Split({" "c, vbTab, vbCrLf, vbCr, vbLf, "#"c}, 2, StringSplitOptions.None).First
                            If Not Provinces.ContainsKey(Province) Then
                                Provinces.Add(Province, Title)
                            Else
                                Provinces(Province) = Title
                            End If
                        End If
                    End With
                End While
                If Not Data.Count = 0 Then
                    LandedTitlesData(Title) = Data
                End If
            End Using
        Next

        Dim NamedColours As New Dictionary(Of String, String)
        Dim NamedColourFiles As List(Of String) = Directory.GetFiles(GameDir & "\common\named_colors").ToList
        For Each ModDir In ModPaths
            Dim NamedColoursPath As String = Path.Combine(ModDir, "common\named_colors")
            If Directory.Exists(NamedColoursPath) Then
                NamedColourFiles.AddRange(Directory.GetFiles(NamedColoursPath).ToList)
            End If
        Next

        For Each ColourFile In NamedColourFiles
            Dim Text As String = File.ReadAllText(ColourFile)
            Dim Index As Integer = Text.IndexOf("{"c, Text.IndexOf("colors") + 6) + 1
            Text = Text.Substring(Index, Text.LastIndexOf("}"c) - Index)
            Do
                Index = Text.IndexOf("="c)
                Dim Name As String = Text.Substring(0, Index).TrimEnd.Split({" "c, vbTab, vbCrLf, vbCr, vbLf}, StringSplitOptions.None).Last
                Index = Text.IndexOf("{"c) + 1
                Dim Colour As String = Text.Substring(Index, Text.IndexOf("}"c, Index) - Index).Trim
                Index = Text.IndexOf(Name) + Name.Length
                If Not Text.Substring(Index, Text.IndexOf("}"c, Index) - Index).Contains("hsv") Then
                    If Not NamedColours.ContainsKey(Name) Then
                        NamedColours.Add(Name, Colour)
                    Else
                        NamedColours(Name) = Colour
                    End If
                Else
                    If Not NamedColours.ContainsKey(Name) Then
                        NamedColours.Add(Name, "255,255,255")
                    Else
                        NamedColours(Name) = "255,255,255"
                    End If
                End If
                Index = Text.IndexOf("}"c, Text.IndexOf(Colour)) + 1
                Text = Text.Substring(Index)
            Loop While Text.Contains("="c)
        Next

        Dim RawProvinces As List(Of String)
        Dim BaronyLieges As New Dictionary(Of Integer, Integer)
        Dim KeyOfTitle As New Dictionary(Of String, Integer)

        Using SR As New StreamReader(GS)
            Dim Delimitor, Splitter, LineData As String
            Dim SB As New StringBuilder()

            'Raw province data for Buildings Data.

            Delimitor = "provinces={"
            Do

            Loop While Not SR.ReadLine.StartsWith(Delimitor)
            Delimitor = "landed_titles={"
            Do
                LineData = SR.ReadLine
                SB.AppendLine(LineData)
            Loop While Not LineData.StartsWith(Delimitor)
            Splitter = vbCrLf & vbTab & "}"c & vbCrLf
            RawProvinces = SB.ToString.Split(Splitter).ToList
            RawProvinces.RemoveAt(RawProvinces.Count - 1)
            Dim RawSpecialBuildings As New Dictionary(Of Integer, String)
            Dim SpecialBuildingsList As New Dictionary(Of String, String)
            For Each Block In RawProvinces
                With Block
                    If .Contains("special_building") Then
                        Dim Province As Integer = .Substring(0, .IndexOf("="c)).TrimStart

                        Dim IsBuilt As Boolean = .Contains("special_building={")
                        Dim StartIndex As Integer = .IndexOf("type=", .IndexOf("special_building")) + 6
                        Dim SpecialBuilding As String = .Substring(StartIndex, .IndexOf(Chr(34), StartIndex) - StartIndex)
                        If Not SpecialBuildingsList.ContainsKey(SpecialBuilding) Then
                            Dim SpecialBuildingName As String = SpecialBuilding.GetLocalisation("building_")
                            SpecialBuildingsList.Add(SpecialBuilding, SpecialBuildingName)
                            SpecialBuilding = SpecialBuildingName
                        Else
                            SpecialBuilding = SpecialBuildingsList(SpecialBuilding)
                        End If
                        If Not IsBuilt Then
                            SpecialBuilding &= " (unbuilt)"
                        End If
                        RawSpecialBuildings.Add(Province, SpecialBuilding)
                    End If
                End With
            Next

            'Landed titles entries title ids and liege and vassal data.

            SB.Clear()
            Delimitor = vbTab & "landed_titles={"
            Do

            Loop While Not SR.ReadLine.StartsWith(Delimitor)
            Delimitor = vbTab & "index="
            Do
                LineData = SR.ReadLine
                SB.AppendLine(LineData)
            Loop While Not LineData.StartsWith(Delimitor)
            Splitter = vbCrLf & "}"c & vbCrLf
            Dim RawTitles As List(Of String) = SB.ToString.Split({vbCrLf & "}"c & vbCrLf, "none" & vbCrLf}, StringSplitOptions.None).ToList
            RawTitles.RemoveAt(RawTitles.Count - 1)
            RawTitles.RemoveAll(Function(x) x.Contains("key=""x_") OrElse x.EndsWith("="c))

            Dim TitleTypes As New Dictionary(Of Integer, Byte)
            Dim RawCapitals As New Dictionary(Of Integer, Integer)
            Dim RawVassals As New Dictionary(Of Integer, List(Of Integer))
            Dim RawLieges As New Dictionary(Of Integer, Dictionary(Of Byte, Integer))
            Dim DuchyLieges, CountyLieges As New Dictionary(Of Integer, Integer)
            Dim KeyOfCounty As New Dictionary(Of String, Integer)
            Dim TitleSplit() As String
            For Each Block In RawTitles
                TitleSplit = Block.Split("={", 2)
                Dim Index As String = TitleSplit(0)
                With TitleSplit(1)

                    'Collect title id

                    Dim StartIndex As Integer = .IndexOf("key=") + 5
                    Dim Title As String = .Substring(StartIndex, .IndexOf(Chr(34), StartIndex) - StartIndex)
                    Dim IsBarony As Boolean = False
                    Dim IsEmpire As Boolean = False
                    Titles.Add(Index, Title)
                    KeyOfTitle.Add(Title, Index)
                    If Title.StartsWith("b"c) Then
                        TitleTypes.Add(Index, b)
                        IsBarony = True
                    ElseIf Title.StartsWith("c"c) Then
                        TitleTypes.Add(Index, c)
                        KeyOfCounty.Add(Title, Index)
                    ElseIf Title.StartsWith("d"c) Then
                        TitleTypes.Add(Index, d)
                    ElseIf Title.StartsWith("k"c) Then
                        TitleTypes.Add(Index, k)
                    Else
                        TitleTypes.Add(Index, e)
                        IsEmpire = True
                    End If

                    'Collect capital data

                    If Not IsBarony Then
                        StartIndex = .IndexOf("capital=")
                        If Not StartIndex = -1 Then
                            StartIndex += 8
                            Dim Capital As Integer = .Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex)
                            RawCapitals.Add(Index, Capital)
                        Else
                            RawCapitals.Add(Index, -1)
                        End If
                    End If

                    'Collect raw vassal data
                    If Not IsBarony Then
                        If .Contains("de_jure_vassals={") Then
                            StartIndex = .IndexOf("de_jure_vassals={") + 17
                            Dim VassalList As String() = .Substring(StartIndex, .IndexOf("}"c, StartIndex) - StartIndex).Trim.Split
                            RawVassals.Add(Index, VassalList.ToList.ConvertAll(Of Integer)(Function(x) x))
                        Else
                            RawVassals.Add(Index, New List(Of Integer))
                        End If
                    End If

                    'Collect liege data

                    If Not IsEmpire Then
                        If .Contains("de_jure_liege=") Then
                            StartIndex = .IndexOf("de_jure_liege=") + 14
                            If TitleTypes(Index) = k Then
                                Dim Liege As Integer = .Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex)
                                RawLieges.Add(Index, New Dictionary(Of Byte, Integer) From {{e, Liege}})
                            ElseIf TitleTypes(Index) = d Then
                                Dim Liege As Integer = .Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex)
                                If RawLieges.ContainsKey(Liege) Then
                                    Dim LiegesList As New Dictionary(Of Byte, Integer)(RawLieges(Liege)) From {{TitleTypes(Liege), Liege}}
                                    RawLieges.Add(Index, LiegesList)
                                Else
                                    DuchyLieges.Add(Index, Liege)
                                End If
                            ElseIf TitleTypes(Index) = c Then
                                Dim Liege As Integer = .Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex)
                                If RawLieges.ContainsKey(Liege) Then
                                    Dim LiegesList As New Dictionary(Of Byte, Integer)(RawLieges(Liege)) From {{TitleTypes(Liege), Liege}}
                                    RawLieges.Add(Index, LiegesList)
                                Else
                                    CountyLieges.Add(Index, Liege)
                                End If
                            ElseIf TitleTypes(Index) = b Then
                                Dim Liege As Integer = .Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex)
                                BaronyLieges.Add(Index, Liege)
                            End If
                        Else
                            RawLieges.Add(Index, New Dictionary(Of Byte, Integer))
                        End If
                    End If
                End With
            Next

            'In case the lieges were collected after the vassal, add them in after the main parsing

            For Each Duchy In DuchyLieges
                With Duchy
                    Dim LiegesList As New Dictionary(Of Byte, Integer)(RawLieges(.Value)) From {{TitleTypes(.Value), .Value}}
                    RawLieges.Add(.Key, LiegesList)
                End With
            Next
            For Each County In CountyLieges
                With County
                    Dim LiegesList As New Dictionary(Of Byte, Integer)(RawLieges(.Value)) From {{TitleTypes(.Value), .Value}}
                    RawLieges.Add(.Key, LiegesList)
                End With
            Next

            'Combining altnames and colours from the earlier parsed landed titles file(s) with the titles collected from the save game.

            For Each Item In Titles
                With Item
                    Dim Name As String = .Value.GetLocalisation
                    If LandedTitlesData.ContainsKey(.Value) Then
                        Dim AltNames As List(Of String) = LandedTitlesData(.Value).FindAll(Function(x) x.StartsWith("cn_"))
                        If Not AltNames.Count = 0 Then
                            AltNames = AltNames.ConvertAll(Function(x) x.GetLocalisation)
                            Name = String.Concat({Name, " "c, "("c, String.Join("/"c, AltNames), ")"c})
                        End If

                        If Not TitleTypes(.Key) = b Then
                            Dim Colour As String() = {LandedTitlesData(.Value).Find(Function(x) x.Contains("color"))}

                            If Colour(0) IsNot Nothing Then
                                If Not Colour(0).Contains("{"c) Then
                                    If Colour(0).Contains(Chr(34)) Then
                                        Dim StartIndex As Integer = Colour(0).IndexOf(Chr(34)) + 1
                                        Colour(0) = Colour(0).Substring(StartIndex, Colour(0).IndexOf(Chr(34), StartIndex) - StartIndex)
                                    Else
                                        Colour(0) = Colour(0).Split("="c, 2).Last.Split({vbCrLf, vbCr, vbLf, vbTab}, StringSplitOptions.None).First.Trim
                                    End If
                                    If NamedColours.ContainsKey(Colour(0)) Then
                                        Colours.Add(.Key, NamedColours(Colour(0)))
                                    Else
                                        Colours.Add(.Key, "255,255,255")
                                        Console.WriteLine($"No colour found for {Titles(.Key)}")
                                    End If
                                Else
                                    Dim StartIndex As Integer = Colour(0).IndexOf("{"c) + 1
                                    Colour(0) = Colour(0).Substring(StartIndex, Colour(0).IndexOf("}"c, StartIndex) - StartIndex).Trim
                                    If Colour(0).Contains("."c) Then
                                        Colour = Colour(0).Split
                                        For Count = 0 To 2
                                            With Colour(Count)
                                                If .Contains("."c) Then
                                                    Colour(Count) = Colour(Count) * 256
                                                    If Colour(Count) >= 256 Then
                                                        Colour(Count) = 255
                                                    End If
                                                    If .Contains("."c) Then
                                                        Colour(Count) = .Remove(.IndexOf("."c))
                                                    End If
                                                End If
                                            End With
                                        Next
                                    Else
                                        Colour = Colour(0).Split
                                    End If
                                    Colours.Add(.Key, String.Join(","c, Colour))
                                End If
                            Else
                                Colours.Add(.Key, "255,255,255")
                                Console.WriteLine($"No colour found for {Titles(.Key)}")
                            End If
                        End If
                    End If
                    Names.Add(.Key, Name)
                End With
            Next

            'Getting capital names

            For Each Capital In RawCapitals
                With Capital
                    If Not .Value = -1 Then
                        Capitals.Add(.Key, Names(.Value))
                    Else
                        Capitals.Add(.Key, "")
                    End If
                End With
            Next

            'Finally parse the lieges into the final format necessary for output.

            For Each Title In RawLieges
                With Title
                    Dim TitleType As Byte = TitleTypes(.Key)
                    If RawLieges(.Key).ContainsKey(e) Then
                        Lieges(e).Add(.Key, Names(RawLieges(.Key)(e)))
                    Else
                        Lieges(e).Add(.Key, "")
                    End If
                    If Not TitleType = k AndAlso RawLieges(.Key).ContainsKey(k) Then
                        Lieges(k).Add(.Key, Names(RawLieges(.Key)(k)))
                    Else
                        Lieges(k).Add(.Key, "")
                    End If
                    If Not TitleType = k AndAlso Not TitleType = d Then
                        Lieges(d).Add(.Key, Names(RawLieges(.Key)(d)))
                    End If
                End With
            Next

            'Sort the vassals

            For Each Title In RawVassals
                With Title
                    Dim T As Byte = TitleTypes(.Key)
                    Dim Count As Integer = 0

                    Do While Count < .Value.Count
                        If .Value.Count >= 20000 Then 'Stopping an infinite loop in case of a bug.
                            Console.WriteLine("An error was encountered where " & Titles(.Key) & " had over 20000 vassals for some reason. Please report this and the mod name to Reciter of Poetry at the CK3 Mod Co-op Discord server. Press any key to exit.")
                            Console.ReadKey(True)
                            Environment.Exit(0)
                        End If
                        Dim Vassal As Integer = .Value(Count)
                        Dim SubVassals As New List(Of Integer)
                        If Not TitleTypes(Vassal) = b Then
                            SubVassals = RawVassals(Vassal)
                        End If
                        If Not SubVassals.Count = 0 Then
                            .Value.AddRange(SubVassals)
                        End If
                        Count += 1
                    Loop

                    If Not .Value.Count = 0 AndAlso Not T = b Then
                        Dim VassalTypes As New List(Of Byte) From {b}
                        If T = e Then
                            VassalTypes.AddRange({c, d, k})
                        ElseIf T = k Then
                            VassalTypes.AddRange({c, d})
                        ElseIf T = d Then
                            VassalTypes.Add(c)
                        ElseIf T = c Then
                            Counties.Add(.Key)
                        End If
                        For Each V In VassalTypes
                            Dim TitleVassals As List(Of String) = .Value.FindAll(Function(x) TitleTypes(x) = V).ConvertAll(Function(x) Names(x))
                            If Not TitleVassals.Count = 0 Then
                                Vassals(V).Add(.Key, TitleVassals)
                                VassalCounts(V).Add(.Key, TitleVassals.Count)
                                If V = c Then
                                    DeJure(T).Add(.Key)
                                End If
                            Else
                                If Not V = c Then
                                    Vassals(V).Add(.Key, TitleVassals)
                                    VassalCounts(V).Add(.Key, 0)
                                Else
                                    Titular(T).Add(.Key)
                                End If
                            End If
                        Next
                    Else
                        Titular(T).Add(.Key)
                    End If
                End With
            Next

            For Each Item In Counties
                Dim Lieges As List(Of Integer) = RawLieges(Item).Values.ToList
                For Each Liege In Lieges
                    If Not LargestCounty.ContainsKey(Liege) Then
                        LargestCounty.Add(Liege, VassalCounts(b)(Item))
                    Else
                        If LargestCounty(Liege) < VassalCounts(b)(Item) Then
                            LargestCounty(Liege) = VassalCounts(b)(Item)
                        End If
                    End If
                Next
            Next

            'Contains faith indices and ids.

            Dim FaithsList As New Dictionary(Of Integer, String)

            SB.Clear()
            Delimitor = vbTab & "faiths={"
            Do

            Loop While Not SR.ReadLine.StartsWith(Delimitor)
            Delimitor = vbTab & "great_holy_wars={"
            Do
                LineData = SR.ReadLine
                SB.AppendLine(LineData)
            Loop While Not LineData.StartsWith(Delimitor)
            Splitter = vbCrLf & vbTab & vbTab & "}"c & vbCrLf
            Dim RawFaiths As List(Of String) = SB.ToString.Split(Splitter).ToList
            RawFaiths.RemoveAt(RawFaiths.Count - 1)

            For Each Block In RawFaiths
                With Block
                    Dim Index As Integer = .Substring(0, .IndexOf("="c)).TrimStart
                    Dim StartIndex As Integer = .IndexOf("tag=") + 5
                    Dim Faith As String = .Substring(StartIndex, .IndexOf(Chr(34), StartIndex) - StartIndex).GetLocalisation
                    FaithsList.Add(Index, Faith)
                End With
            Next

            'Contains counties and lists their development, and culture and faith indices.

            SB.Clear()
            Delimitor = "county_manager={"
            Do

            Loop While Not SR.ReadLine.StartsWith(Delimitor)
            SR.ReadLine()
            Delimitor = vbTab & "monthly_increase={"
            Do
                LineData = SR.ReadLine
                SB.AppendLine(LineData)
            Loop While Not LineData.StartsWith(Delimitor)
            Splitter = vbCrLf & vbTab & vbTab & "}"c & vbCrLf
            Dim RawCounties As List(Of String) = SB.ToString.Split(Splitter).ToList
            RawCounties.RemoveAt(RawCounties.Count - 1)

            'Contains culture indices and ids.

            Dim CulturesList As New Dictionary(Of Integer, String)

            SB.Clear()
            Delimitor = "culture_manager={"
            Do

            Loop While Not SR.ReadLine.StartsWith(Delimitor)
            SR.ReadLine()
            Delimitor = vbTab & "template_cultures={"
            Do
                LineData = SR.ReadLine
                SB.AppendLine(LineData)
            Loop While Not LineData.StartsWith(Delimitor)
            Splitter = vbCrLf & vbTab & vbTab & "}"c & vbCrLf
            Dim RawCultures As List(Of String) = SB.ToString.Split(Splitter).ToList
            RawCultures.RemoveAt(RawCultures.Count - 1)

            For Each Block In RawCultures
                With Block
                    Dim Index As Integer = .Substring(0, .IndexOf("="c)).TrimStart
                    Dim StartIndex As Integer = .IndexOf("name=") + 6
                    Dim Culture As String = .Substring(StartIndex, .IndexOf(Chr(34), StartIndex) - StartIndex).GetLocalisation
                    CulturesList.Add(Index, Culture)
                End With
            Next

            'Collecting county faiths and cultures.

            Dim ModifiersList As New Dictionary(Of String, String)
            For Each Block In RawCounties
                With Block
                    Dim County As Integer = KeyOfCounty(.Substring(0, .IndexOf("="c)).TrimStart)

                    Dim StartIndex As Integer = .IndexOf("development=") + 12
                    Dim Development As Integer = .Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex)
                    CountyDev.Add(County, Development)

                    StartIndex = .IndexOf("culture=") + 8
                    Dim Culture As String = CulturesList(.Substring(StartIndex, .IndexOf(vbCr, StartIndex) - StartIndex))
                    Cultures.Add(County, Culture)
                    StartIndex = .IndexOf("faith=") + 6
                    Dim EndIndex As Integer = .IndexOf(vbCr, StartIndex)
                    If EndIndex = -1 Then
                        EndIndex = .Length
                    End If
                    EndIndex -= StartIndex
                    Dim Faith As String = FaithsList(.Substring(StartIndex, EndIndex))
                    Faiths.Add(County, Faith)

                    If .Contains("modifier=""") Then
                        StartIndex = 0
                        Dim CountyModifiers As New Dictionary(Of String, Integer)
                        Do While Not .IndexOf("modifier=""", StartIndex) = -1
                            StartIndex = .IndexOf("modifier=""", StartIndex) + 10
                            Dim Modifier As String = .Substring(StartIndex, .IndexOf(Chr(34), StartIndex) - StartIndex)
                            If Not ModifiersList.ContainsKey(Modifier) Then
                                Dim ModifierName As String = Modifier.GetLocalisation
                                ModifiersList.Add(Modifier, ModifierName)
                                Modifier = ModifierName
                            Else
                                Modifier = ModifiersList(Modifier)
                            End If
                            CountyModifiers.Add(Modifier, 1)
                        Loop
                        Modifiers.Add(County, CountyModifiers)
                    End If
                End With
            Next

            Dim CountyKeys As List(Of Integer) = CountyDev.Keys.ToList
            For Each Item In CountyKeys
                Dim LiegesKeys As List(Of Integer) = RawLieges(Item).Values.ToList
                Dim Development As Integer = CountyDev(Item)
                For Each Liege In LiegesKeys
                    If Not TotalDev.ContainsKey(Liege) Then
                        TotalDev.Add(Liege, Development)
                        HighestDev.Add(Liege, Development)
                    Else
                        TotalDev(Liege) += Development
                        If HighestDev(Liege) < Development Then
                            HighestDev(Liege) = Development
                        End If
                    End If
                Next
            Next

            CountyKeys = Modifiers.Keys.ToList
            For Each Item In CountyKeys
                Dim LiegesKeys As List(Of Integer) = RawLieges(Item).Values.ToList

                For Each Liege In LiegesKeys
                    If Not Modifiers.ContainsKey(Liege) Then
                        Modifiers.Add(Liege, Modifiers(Item))
                    Else
                        With Modifiers(Liege)
                            For Each Modifier In Modifiers(Item)
                                If Not .ContainsKey(Modifier.Key) Then
                                    .Add(Modifier.Key, Modifier.Value)
                                Else
                                    Modifiers(Liege)(Modifier.Key) += Modifier.Value
                                End If
                            Next
                        End With
                    End If
                Next
            Next

            For Each Item In RawSpecialBuildings
                Dim County As Integer = BaronyLieges(KeyOfTitle(Provinces(Item.Key)))
                If Not SpecialBuildings.ContainsKey(County) Then
                    SpecialBuildings.Add(County, New Dictionary(Of String, Integer) From {{Item.Value, 1}})
                Else
                    If Not SpecialBuildings(County).ContainsKey(Item.Value) Then
                        SpecialBuildings(County).Add(Item.Value, 1)
                    Else
                        SpecialBuildings(County)(Item.Value) += 1
                    End If
                End If
            Next

            CountyKeys = SpecialBuildings.Keys.ToList
            For Each Item In CountyKeys
                Dim LiegesKeys As List(Of Integer) = RawLieges(Item).Values.ToList

                For Each Liege In LiegesKeys
                    If Not SpecialBuildings.ContainsKey(Liege) Then
                        SpecialBuildings.Add(Liege, SpecialBuildings(Item))
                    Else
                        With SpecialBuildings(Liege)
                            For Each SpecialBuilding In SpecialBuildings(Item)
                                If Not .ContainsKey(SpecialBuilding.Key) Then
                                    .Add(SpecialBuilding.Key, SpecialBuilding.Value)
                                Else
                                    SpecialBuildings(Liege)(SpecialBuilding.Key) += SpecialBuilding.Value
                                End If
                            Next
                        End With
                    End If
                Next
            Next
        End Using

        Dim EndTime As DateTime = Now
        Debug.Print(EndTime.Subtract(StartTime).TotalSeconds.ToString)

        Console.WriteLine()
        Console.WriteLine("Enter output code or path to file with codes.")
        Dim Input As String = Console.ReadLine
        If Not Path.HasExtension(Input) Then
            Dim Response As Boolean = True
            Do While Response = True

                WriteToOutput(Input)

                Console.WriteLine("The output file has been deposited in your desktop. Do you wish to create another table?")
                Console.WriteLine("Press Y to continue or anything else to exit.")
                Dim YorN As ConsoleKey = Console.ReadKey.Key
                If Not YorN = ConsoleKey.Y Then
                    Response = False
                Else
                    Console.WriteLine()
                    Console.WriteLine("Enter output code.")
                    Input = Console.ReadLine
                End If
            Loop
        Else
            Dim Codes() As String = File.ReadAllLines(Input)
            For Each Input In Codes
                WriteToOutput(Input)
            Next
            Console.WriteLine("The output files have been deposited in your desktop. Press any key to exit.")
            Console.ReadKey(True)
        End If

        'Extension for Godherja Metropoleis. This isn't favouritism or bias, I'm creating the wiki lists for the Godherja wiki right now and I'm not going to goddamn do these by hand.
        If ModNames.Exists(Function(x) x.Contains("Godherja")) Then
            Dim OutputData As New Dictionary(Of String, Integer)
            Dim MetropoleisList As New Dictionary(Of String, String)
            Dim RawMetropoleis As New List(Of Integer)
            For Each Block In RawProvinces
                With Block
                    If .Contains("district_holding") Then
                        Dim Province As Integer = .Substring(0, .IndexOf("="c)).TrimStart
                        RawMetropoleis.Add(Province)
                    End If
                End With
            Next

            For Each Item In RawMetropoleis
                Dim CountyKey As Integer = BaronyLieges(KeyOfTitle(Provinces(Item)))
                Dim County As String = Names(CountyKey)
                If Not OutputData.ContainsKey(County) Then
                    OutputData.Add(County, 1)
                Else
                    OutputData(County) += 1
                End If
                Dim Duchy As String = Lieges(d)(CountyKey)
                If Not OutputData.ContainsKey(Duchy) Then
                    OutputData.Add(Duchy, 1)
                Else
                    OutputData(Duchy) += 1
                End If
            Next

            Dim TextFiles As New List(Of String) From {Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{OutputName} De Jure Duchies.txt"), Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{OutputName} Counties.txt")}
            For Each TextFile In TextFiles
                If File.Exists(TextFile) Then
                    Dim Text As List(Of String) = File.ReadAllText(TextFile).Split("|-").ToList
                    Text(Text.Count - 1) = $"{Text(Text.Count - 1).Substring(0, Text(Text.Count - 1).LastIndexOf("|}"))}"
                    If Text(0).Contains("Baronies") Then
                        Dim ColumnCount As Integer = Text(0).Substring(0, Text(0).IndexOf("Baronies")).Split("!"c).Length + 4
                        Text(0) = Text(0).Insert(Text(0).IndexOf("Baronies") + 10, " Metropoleis!!")
                        For Count = 1 To Text.Count - 1 'Put - 2 here before for some reason.
                            Dim Title As String = Text(Count).Split("||")(1)
                            Dim Insertion As String = ""
                            If OutputData.ContainsKey(Title) Then
                                Insertion = OutputData(Title)
                            End If
                            Text(Count) = Text(Count).Insert(GetNthIndex(Text(Count), "|"c, ColumnCount) + 1, $"{Insertion}||")
                        Next
                        Text(Text.Count - 1) &= "|}"
                        File.WriteAllText(TextFile, String.Join("|-", Text))
                    End If
                End If
            Next
        End If
    End Sub
    'SetUp and Functions
    Sub SetUp()
        'TODO: Remove read and write time checks and make sure it always extracts.
        If Directory.GetFiles(BaseDir).ToList.Exists(Function(x) x.EndsWith(".ck3")) Then
            Dim SaveFile As String = Directory.GetFiles(BaseDir).ToList.FindAll(Function(x) x.EndsWith(".ck3")).OrderByDescending(Function(x) New FileInfo(x).CreationTime).First()
            If File.Exists(GS) AndAlso DateTime.Compare(File.GetCreationTime(SaveFile), File.GetCreationTime(GS)) < 0 Then
                Dim IsIdentical = True
                Using GSR As New StreamReader(GS)
                    Using SFR As New StreamReader(SaveFile)
                        SFR.ReadLine()
                        Dim LineData As String = ""
                        Dim Closer As String = vbTab & "ironman=no"
                        Do While Not LineData = Closer
                            LineData = SFR.ReadLine
                            If Not LineData = GSR.ReadLine Then
                                IsIdentical = False
                            End If
                        Loop
                    End Using
                End Using
                If Not IsIdentical Then
                    ExtractSaveFile(SaveFile)
                End If
            Else
                ExtractSaveFile(SaveFile)
            End If
        Else
            Dim SaveFileFound As Boolean = False
            Do
                Console.WriteLine("Please place the save file you wish to parse in the root directory of this application. Press any key to continue once you have done so.")
                Console.ReadKey(True)
                If Directory.GetFiles(BaseDir).ToList.Exists(Function(x) x.EndsWith(".ck3")) Then
                    SaveFileFound = True
                    Dim SaveFile As String = Directory.GetFiles(BaseDir).ToList.FindAll(Function(x) x.EndsWith(".ck3")).OrderByDescending(Function(x) New FileInfo(x).LastWriteTime).First()
                    If File.Exists(GS) AndAlso DateTime.Compare(File.GetLastWriteTime(SaveFile), File.GetLastWriteTime(GS)) < 0 Then
                        Dim IsIdentical = True
                        Using GSR As New StreamReader(GS)
                            Using SFR As New StreamReader(SaveFile)
                                SFR.ReadLine()
                                Dim LineData As String = ""
                                Dim Closer As String = vbTab & "ironman=no"
                                Do While Not LineData = Closer
                                    LineData = SFR.ReadLine
                                    If Not LineData = GSR.ReadLine Then
                                        IsIdentical = False
                                    End If
                                Loop
                            End Using
                        End Using
                        If Not IsIdentical Then
                            ExtractSaveFile(SaveFile)
                        End If
                    Else
                        If File.Exists(GS) Then
                            File.Delete(GS)
                        End If
                        ExtractSaveFile(SaveFile)
                    End If
                End If
            Loop While SaveFileFound = False
        End If

        Using SR As New StreamReader(GS)
            Dim LineData As String
            While Not SR.EndOfStream
                LineData = SR.ReadLine
                If LineData.TrimStart.StartsWith("mods=") Then
                    Dim ModsList As List(Of String) = LineData.Split("mods={ ")(1).Split(" }", 2)(0).Replace(Chr(34), "").Replace("mod/", "").Split.ToList
                    If OperatingSystem.IsWindows OrElse OperatingSystem.IsMacOS Then
                        ModsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                    ElseIf OperatingSystem.IsLinux Then
                        ModsFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData)
                    End If
                    If Not Path.EndsInDirectorySeparator(ModsFolder) Then
                        ModsFolder &= Path.DirectorySeparatorChar
                    End If
                    ModsFolder &= "Paradox Interactive/Crusader Kings III/mod/"
                    For Each Item In ModsList
                        Dim ModFile As List(Of String) = File.ReadAllLines(ModsFolder & Item).ToList
                        Dim Name As String = ModFile.Find(Function(x) x.StartsWith("name=")).Split(Chr(34), 3)(1)
                        OutputName &= Name
                        If ModFile.Exists(Function(x) x.StartsWith("path=")) Then
                            Dim ModPath As String = ModFile.Find(Function(x) x.StartsWith("path=")).Split(Chr(34), 3)(1)
                            Mods.Add(Item)
                            ModPaths.Add(ModPath)
                            ModNames.Add(Name)
                        ElseIf ModFile.Exists(Function(x) x.StartsWith("archive=")) Then
                            Dim ModPath As String = ModFile.Find(Function(x) x.StartsWith("archive=")).Split(Chr(34), 3)(1)
                            Mods.Add(Item)
                            Dim TempPath As String = Path.GetTempPath & "/CK3Tools/" & Item
                            Directory.CreateDirectory(TempPath)
                            ModPaths.Add(TempPath)
                            ModNames.Add(Name)
                            ExtractModArchive(ModPath, TempPath)
                        Else
                            Console.WriteLine($"Cannot find the directory of the {Name} mod.")
                        End If
                    Next
                ElseIf LineData.StartsWith("date=") Then
                    SaveDate = LineData.Split("="c, 2).Last
                ElseIf LineData.StartsWith("variables=") Then
                    If Mods.Count = 0 Then
                        OutputName = "BaseGame"
                    End If
                    Exit While
                End If
            End While
        End Using

        OutputName = String.Concat(OutputName.Split(Path.GetInvalidFileNameChars)).Replace(" ", "")

        SetGameDir()

    End Sub
    Sub ExtractSaveFile(SaveFile As String)
        Dim SevenZip As Process
        If OperatingSystem.IsWindows Then
            If Not File.Exists(BaseDir & "/7za.exe") Then
                Console.WriteLine("The external tool 7zip has not been copied correctly. Please extract the 'gamestate' file from the .ck3 save file and place it in the root directory of this app. Then press any key to continue.")
                Console.ReadKey(True)
                If Not File.Exists(GS) Then
                    Console.WriteLine("Gamestate file not found.")
                End If
            Else
                SevenZip = Process.Start("cmd", "/c" & $"7za.exe  x -aoa -r ""{SaveFile}"" -o""{BaseDir}"" *gamestate*")
                SevenZip.WaitForExit()
                Console.Clear()
            End If
        ElseIf OperatingSystem.IsMacOS OrElse OperatingSystem.IsLinux Then
            Do While Not File.Exists(GS)
                Console.WriteLine("This app does not properly support MacOS or Linux currently. Please extract the 'gamestate' file from the .ck3 save file and place it in the root directory of this app. Then press any key to continue.")
                Console.ReadKey(True)
            Loop
        End If
        File.SetCreationTime(GS, Now)
    End Sub
    Sub ExtractModArchive(ArchivePath As String, TempFolder As String)
        Dim ModArchive As ZipArchive = ZipFile.OpenRead(ArchivePath)
        Dim PathsToExtract As New List(Of String) From {
            "localization/english",
            "common/landed_titles"}
        For Each ModPath In PathsToExtract
            ModArchive.ExtractFolder(ModPath, TempFolder)
        Next
    End Sub
    <Extension()>
    Sub ExtractFolder(ByRef Archive As ZipArchive, Folder As String, OutputFolder As String)
        Dim Entries As List(Of ZipArchiveEntry) = Archive.Entries.ToList.FindAll(Function(x) x.FullName.StartsWith(Folder))
        For Each Entry In Entries
            Dim EntryFullName As String = Path.Combine(OutputFolder, Entry.FullName)
            Dim EntryName As String = Path.GetFileName(EntryFullName)
            If Not String.IsNullOrEmpty(EntryName) Then
                Dim EntryPath As String = Path.GetDirectoryName(EntryFullName)
                If Not Directory.Exists(EntryPath) Then
                    Directory.CreateDirectory(EntryPath)
                End If
                Entry.ExtractToFile(EntryFullName)
            End If
        Next
    End Sub
    Sub SetGameDir()
        If ModPaths.Exists(Function(x) x.Contains("steamapps")) Then
            GameDir = ModPaths.Find(Function(x) x.Contains("steamapps")).Split("steamapps", 2).First.Split(":"c, 2).Last & "steamapps\common\Crusader Kings III\game\"
        ElseIf File.Exists(Path.GetTempPath() & Path.DirectorySeparatorChar & "CK3Tools.txt") Then
            Using SR As New StreamReader(Config)
                GameDir = SR.ReadLine
                If Not Path.EndsInDirectorySeparator(GameDir) Then
                    GameDir &= Path.DirectorySeparatorChar
                End If
                If Directory.GetDirectories(GameDir, SearchOption.TopDirectoryOnly).ToList.Exists(Function(x) x.Contains("binaries")) Then
                    Dim Binaries As String = Path.Combine(GameDir, "binaries")
                    If Not Directory.GetFiles(GameDir, SearchOption.AllDirectories).ToList.Contains(Path.Combine(Binaries & "ck3.exe")) Then
                        GameDirPrompt()
                    End If
                End If
            End Using
        Else
            GameDirPrompt()
        End If
    End Sub
    Sub GameDirPrompt()
        Dim DirFound As Boolean = False
        Do
            Console.WriteLine("Please enter your Crusader Kings 3 installation's root directory.")
            GameDir = Console.ReadLine
            If Not Path.EndsInDirectorySeparator(GameDir) Then
                GameDir &= Path.DirectorySeparatorChar
            End If
            If Directory.GetDirectories(GameDir, SearchOption.TopDirectoryOnly).ToList.Exists(Function(x) x.Contains("binaries")) Then
                Dim Binaries As String = Path.Combine(GameDir, "binaries")
                If Directory.GetFiles(GameDir, SearchOption.AllDirectories).ToList.Contains(Path.Combine(Binaries & "ck3.exe")) Then
                    DirFound = True
                End If
            End If
        Loop While DirFound = False

        Using SW As New StreamWriter(Config, False)
            SW.WriteLine(GameDir)
        End Using
    End Sub
    'CollectLocalisations and Functions
    Sub CollectLocalisations()
        Dim RawGameConceptLocalisations As New Dictionary(Of String, String)

        Dim LocalisationFiles As List(Of String) = Directory.GetFiles(GameDir & "\localization\english", "*.yml", SearchOption.AllDirectories).ToList

        Dim Langs As List(Of String) = Directory.GetDirectories(GameDir & "\localization").ToList.ConvertAll(Function(x) $"{Path.DirectorySeparatorChar}{x.Split({Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar}).Last}")
        Langs.Add(Path.DirectorySeparatorChar & "replace") 'Determine all the languages supported by gase games. Save their folder names and their replace folder names.

        For Each ModPath In ModPaths

            Dim AltDirs As List(Of String) = Directory.EnumerateDirectories(ModPath & "\localization").ToList.FindAll(Function(x) Not Langs.Contains($"{Path.DirectorySeparatorChar}{x.Split({Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar}).Last}"))

            If Directory.Exists(ModPath & "\localization\english") Then
                LocalisationFiles = LocalisationFiles.Concat(Directory.GetFiles(ModPath & "\localization\english", "*.yml", SearchOption.AllDirectories)).ToList
            End If
            If Directory.Exists(ModPath & "\localization\replace\english") Then
                LocalisationFiles = LocalisationFiles.Concat(Directory.GetFiles(ModPath & "\localization\replace\english", "*.yml", SearchOption.AllDirectories)).ToList
            End If
            If Not AltDirs.Count = 0 Then
                For Each AltDir In AltDirs
                    LocalisationFiles = LocalisationFiles.Concat(Directory.GetFiles(AltDir, "*.yml", SearchOption.AllDirectories)).ToList
                Next
            End If
        Next

        For Each Textfile In LocalisationFiles
            SaveLocs(Textfile, RawGameConceptLocalisations)
        Next

        For Each Item In RawGameConceptLocalisations
            GameConceptLocalisations.Add(Item.Key, Item.Value.Split(Chr(34), 2).Last.DeComment.TrimEnd.TrimEnd(Chr(34)).DeFormat.TrimEnd)
        Next
        For Count = 0 To GameConceptLocalisations.Count - 1
            If GameConceptLocalisations.Values(Count).Contains("$"c) Then
                Dim Key As String = GameConceptLocalisations.Keys(Count)
                GameConceptLocalisations(Key) = GameConceptLocalisations.Values(Count).ToString.DeReference
            End If
        Next
    End Sub
    <Extension()>
    Private Sub GetLocalisation(ByRef Code As List(Of String), Optional Prefix As String = "", Optional Suffix As String = "")
        For Count = 0 To Code.Count - 1
            If Not Code(Count).Length = 0 AndAlso Not Code(Count).TrimStart.StartsWith("game_concept") Then
                Dim RawCode As String = String.Concat({Prefix, Code(Count), Suffix}) 'Modify the code if the object id has a suffix in the loc code.
                If SavedLocalisation.Contains(RawCode) Then 'If the locs stored to dictionary contain this loc then...
                    Code(Count) = SavedLocalisation(RawCode)

                    'Process the loc for internal code.

                    If Code(Count).Split(Chr(34)).Last.Contains("#"c) Then
                        Code(Count) = Code(Count).DeComment 'Remove comments if any.
                    End If
                    Code(Count) = Code(Count).Split(Chr(34), 2).Last.TrimEnd.TrimEnd(Chr(34))
                    If Code(Count).Contains("#"c) Then
                        Code(Count) = Code(Count).DeFormat 'Remove style formatting if any.
                    End If
                    If Code(Count).Contains("|E]") OrElse Code(Count).Contains("|e]") Then
                        Code(Count) = Code(Count).DeConcept 'Find the appropriate locs for any game concepts referred.
                    End If
                    If Code(Count).Contains("$"c) Then
                        Code(Count) = Code(Count).DeReference 'Find the appropriate locs for any other locs referred.
                    End If
                ElseIf GameConceptLocalisations.Contains(RawCode) Then 'If game concept locs stored to dictionary contain this loc then...
                    Code(Count) = GameConceptLocalisations(RawCode) 'Get the loc from the game concept dictionary.
                Else 'If the locs stored to memory or the game concept locs stored to memory don't contain this loc then...
                    Code(Count) = RawCode 'Write down the code without any localisation.
                End If
            ElseIf Code(Count).TrimStart.StartsWith("game_concept") AndAlso GameConceptLocalisations.Contains(Code(Count).Split("game_concept_", 2).Last.Split(":"c, 2).First) Then 'If the loc starts with game_concept then look for it in the game concept dictionary.
                Code(Count) = GameConceptLocalisations(Code(Count).Split("game_concept_").Last)
            End If
        Next
    End Sub
    <Extension()>
    Function GetLocalisation(Input As String, Optional Prefix As String = "", Optional Suffix As String = "") As String
        If Not Input.Length = 0 AndAlso Not Input.TrimStart.StartsWith("game_concept") Then
            Dim RawCode As String = Prefix & Input & Suffix 'Modify the code if the object id has a suffix in the loc code.
            If SavedLocalisation.Contains(RawCode) Then 'If the locs stored to dictionary contain this loc then...
                Input = SavedLocalisation(RawCode)

                'Process the loc for internal code.

                If Input.Split(Chr(34)).Last.Contains("#"c) Then
                    Input = Input.DeComment 'Remove comments if any.
                End If
                Input = Input.Split(Chr(34), 2).Last.TrimEnd.TrimEnd(Chr(34))
                If Input.Contains("#"c) Then
                    Input = Input.DeFormat 'Remove style formatting if any.
                End If
                If Input.Contains("|E]") OrElse Input.Contains("|e]") Then
                    Input = Input.DeConcept 'Find the appropriate locs for any game concepts referred.
                End If
                If Input.Contains("$"c) Then
                    Input = Input.DeReference 'Find the appropriate locs for any other locs referred.
                End If
            ElseIf GameConceptLocalisations.Contains(RawCode) Then 'If game concept locs stored to dictionary contain this loc then...
                Input = GameConceptLocalisations(RawCode) 'Get the loc from the game concept dictionary.
            Else 'If the locs stored to memory or the game concept locs stored to memory don't contain this loc then...
                Input = RawCode 'Write down the code without any localisation.
            End If
        ElseIf Input.TrimStart.StartsWith("game_concept") AndAlso GameConceptLocalisations.Contains(Input.Split("game_concept_", 2).Last.Split(":"c, 2).First) Then 'If the loc starts with game_concept then look for it in the game concept dictionary.
            Input = GameConceptLocalisations(Input.Split("game_concept_").Last)
        End If
        Return Input
    End Function
    Sub SaveLocs(TextFile As String, RawGameConceptLocalisations As Dictionary(Of String, String))
        Using SR As New StreamReader(TextFile)
            Dim LineData As String
            While Not SR.EndOfStream
                LineData = SR.ReadLine
                If Not LineData.TrimStart.StartsWith("#"c) AndAlso LineData.Contains(":"c) AndAlso Not LineData.Split(":"c, 2).Last.Length = 0 Then

                    Dim Key As String = LineData.TrimStart.Split(":"c, 2).First
                    Dim Value As String = LineData.Split(":"c, 2).Last.Substring(1).TrimStart

                    If Not SavedLocalisation.Contains(Key) Then
                        SavedLocalisation.Add(Key, Value)
                    Else
                        SavedLocalisation(Key) = Value
                    End If
                End If
                If LineData.TrimStart.StartsWith("game_concept") Then
                    Dim Key As String = LineData.TrimStart.Split(":"c, 2).First.Split("game_concept_", 2).Last
                    Dim Value As String = LineData.Split(":"c, 2).Last.Substring(1).TrimStart

                    If Not RawGameConceptLocalisations.ContainsKey(Key) Then
                        RawGameConceptLocalisations.Add(Key, Value)
                    Else
                        RawGameConceptLocalisations(Key) = Value
                    End If
                End If
            End While
        End Using
    End Sub
    <Extension()>
    Function DeComment(ByVal Input As String) As String
        Dim Output As List(Of String) = Input.Split(Chr(34)).ToList 'Find the boundaries of the actual loc code by splitting it up according to its quotation marks.
        Output(Output.Count - 1) = Output(Output.Count - 1).Split("#"c).First 'Take the last part of the split input, and split it off from the comment.
        Return String.Join(Chr(34), Output).TrimEnd 'Rejoin the input with quotation marks and return it.
    End Function
    <Extension()>
    Function DeConcept(Input As String) As String
        If Input.Contains("|E]") OrElse Input.Contains("|e]") Then
            Dim GameConcepts As New SortedList(Of String, String) 'Collect each game concept contained in string here.
            Do
                Dim GameConcept As String = Input.Split("["c, 2).Last.Split("|"c, 2).First 'Get the game concept object id.
                If Not GameConcepts.ContainsKey(GameConcept) Then 'If it has not already been collected then...
                    Dim ReplaceString As String
                    If GameConceptLocalisations.Contains(GameConcept.ToLower) Then 'Find its loc in the SortedList.
                        ReplaceString = GameConceptLocalisations(GameConcept.ToLower)
                    Else
                        ReplaceString = GameConcept 'If it cannot be found then assign the replace string to be the raw code.
                    End If
                    GameConcepts.Add(GameConcept, ReplaceString) 'Add it to the sortedlist and find the rest of the game concepts in this loc string.
                    Input = Input.Replace($"[{GameConcept}|E]", ReplaceString).Replace($"[{GameConcept}|e]", ReplaceString) 'Remove it from the input string so it is not reparsed into the SortedList.
                Else 'If it has already been collected...
                    'Input = String.Concat(Input.Split({"[", "|E]"}, 3, StringSplitOptions.None)) 'Remove it from the input string.
                    Input = Input.Replace($"[{GameConcept}|E]", GameConcepts(GameConcept)).Replace($"[{GameConcept}|e]", GameConcepts(GameConcept))
                End If
            Loop While Input.Contains("|E]") OrElse Input.Contains("|e]") 'Loop while input loc string contains any non-parsed game concepts.

            Return Input 'Return loc.
        Else
            Return Input 'Redundancy in case a loc was falsely found to contain a game concept.
        End If

    End Function
    <Extension()>
    Function DeReference(Input As String) As String
        If Input.Contains("$"c) Then
            Dim Output As String = Input
            Dim Locs As New List(Of String)
            Do
                If Not Locs.Contains(Input.Split("$", 3)(1)) Then
                    Locs.Add(Input.Split("$", 3)(1))
                End If
                Input = Input.Split("$", 3).Last
            Loop While Input.Contains("$"c)
            Dim Code As List(Of String) = Locs.ToList
            GetLocalisation(Locs)
            For Count = 0 To Code.Count - 1
                Output = Output.Replace($"${Code(Count)}$", Locs(Count))
            Next
            Return Output
        Else
            Return Input
        End If
    End Function
    <Extension()>
    Function DeFormat(ByVal Input As String) As String
        If Input.Contains("#"c) Then
            Dim Output As String = Input
            Do
                Dim FormattedLoc As String = Output.Split("#"c, 2).Last 'Find the styled part of the loc and extract it.
                Dim DeFormatted As String
                With FormattedLoc
                    Dim CloserIndex As Integer
                    If .Contains("#!") Then
                        CloserIndex = .IndexOf("#!") + 2
                        DeFormatted = .Substring(0, CloserIndex - 2).Split(" "c, 2).Last
                    ElseIf .Contains("#"c) Then
                        CloserIndex = .IndexOf("#"c) + 1
                        DeFormatted = .Substring(0, CloserIndex - 1).Split(" "c, 2).Last
                    Else
                        CloserIndex = .Length
                        DeFormatted = FormattedLoc.Split(" "c, 2).Last
                    End If
                    FormattedLoc = String.Concat("#"c, .Substring(0, CloserIndex))
                End With
                'Remove the style code and store it in a string.

                Output = Output.Replace(FormattedLoc, DeFormatted) 'Use replace function to replace the formatted part of the loc with the deformatted string.
            Loop While Output.Contains("#"c) 'Loop if there are more.
            Return Output 'Return when there are no more.
        Else
            Return Input
        End If
    End Function
    <Extension()>
    Function DeNest(Input As String) As List(Of String)
        Dim Output As New List(Of String)
        Input = String.Join(vbCrLf, Input.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).ToList.FindAll(Function(x) Not x.TrimStart.StartsWith("#"c)))
        If Input.Contains("="c) AndAlso Input.Contains("{"c) Then
            Do
                Dim RawCodeID As String = Input.Split("{"c, 2).First.Split("="c, 2).First & "="c 'Get the code id of the object the block is assigned to.
                Input = Input.Substring(RawCodeID.Length - 1) 'Split off the extracted data.
                Dim RawCodeBlock As String = Input
                RawCodeID = RawCodeID.Split({vbCrLf, vbCr, vbLf}, StringSplitOptions.None).Last
                If RawCodeBlock.Split("}"c, 2).First.Contains("{"c) AndAlso RawCodeBlock.Split("{"c, 2).First.Contains("="c) Then
                    RawCodeBlock = RawCodeBlock.Split({vbCrLf, vbCr, vbCrLf}, 2, StringSplitOptions.None).First
                    Output.Add(String.Concat({RawCodeID, RawCodeBlock}))
                Else
                    Do While RawCodeBlock.Split("}"c, 2)(0).Contains("{"c) 'Designate subsidiary objects designated with curly brackets as such by replacing their {} with <>.
                        RawCodeBlock = String.Join(">"c, String.Join("<"c, RawCodeBlock.Split("{"c, 2)).Split("}"c, 2))
                    Loop  'Loop until no more subsidiary objects.
                    'End If
                    RawCodeBlock = RawCodeBlock.Split("}"c)(0).Replace("<"c, "{"c).Replace(">"c, "}"c) & "}"c 'Get the data of this object by splitting it off of the overall code after its own { closing bracket.
                    If RawCodeID.Contains("="c) Then
                        Output.Add(String.Join("{"c, {RawCodeID, RawCodeBlock})) 'Add to List
                    End If
                End If

                If Input.Length > RawCodeBlock.Length Then 'Split off the extracted code block.
                    Input = Input.Substring(RawCodeBlock.Length - 1)
                Else
                    Input = ""
                End If
            Loop While Input.Split("}"c, 2)(0).Contains("{"c) 'Continue to parse the data until no more faiths can be found by looking for a { starting bracket.
        End If
        Return Output
    End Function
    Sub WriteToOutput(Input As String)
        Static ColumnOptions As New Dictionary(Of String, Byte) From {{"t"c, 1}, {"n"c, 2}, {"le", 3}, {"lk", 4}, {"ld", 5}, {"vk", 6}, {"cvk", 7}, {"vd", 8}, {"cvd", 9}, {"vc", 10}, {"cvc", 11}, {"vb", 12}, {"cvb", 13}, {"lc", 14}, {"cpt", 15}, {"clt", 16}, {"f"c, 17}, {"dt", 18}, {"dh", 19}, {"dc", 20}, {"sb", 21}, {"m"c, 22}, {"sbm", 23}}
        Static TitleOptionNames As New Dictionary(Of String, String) From {{"DJE", "De Jure Empires"}, {"TE", "Titular Empires"}, {"DJK", "De Jure Kingdoms"}, {"TK", "Titular Kingdoms"}, {"DJD", "De Jure Duchies"}, {"TD", "Titular Duchies"}, {"C", "Counties"}}
        Static ColumnNames As New Dictionary(Of Byte, String) From {{1, "Title ID"}, {2, "Name"}, {3, "Empire"}, {4, "Kingdom"}, {5, "Duchy"}, {6, "Kingdoms"}, {7, "Kingdoms"}, {8, "Duchies"}, {9, "Duchies"}, {10, "Counties"}, {11, "Counties"}, {12, "Baronies"}, {13, "Baronies"}, {14, "Largest County"}, {15, "Capital"}, {16, "Culture"}, {17, "Faith"}, {18, "Total Dev"}, {19, "Highest Dev"}, {20, "Development"}, {21, "Other"}, {22, "Other"}, {23, "Other"}}

        Dim TitleOption As String
        Dim OutputTitles As New List(Of Integer)
        Dim Columns As List(Of Byte)

        With Input
            TitleOption = .Substring(0, .IndexOf(" "c))
            Select Case TitleOption
                Case "DJE"
                    OutputTitles = DeJure(e)
                Case "TE"
                    OutputTitles = Titular(e)
                Case "DJK"
                    OutputTitles = DeJure(k)
                Case "TK"
                    OutputTitles = Titular(k)
                Case "DJD"
                    OutputTitles = DeJure(d)
                Case "TD"
                    OutputTitles = Titular(d)
                Case "C"
                    OutputTitles = Counties
                Case Else
                    Exit Select
            End Select
            Dim ColumnsCheck As List(Of String) = .Substring(.IndexOf(" "c) + 1).Split(" "c, StringSplitOptions.TrimEntries).ToList
            ColumnsCheck.RemoveAll(Function(x) Not ColumnOptions.ContainsKey(x))
            Columns = ColumnsCheck.ConvertAll(Function(x) ColumnOptions(x))
        End With

        If Not OutputTitles.Count = 0 Then
            Using SW As New StreamWriter(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), $"{OutputName} {TitleOptionNames(TitleOption)}.txt"), False)
                SW.WriteLine("{| class=""wikitable sortable""")
                Dim SB As New StringBuilder
                SB.Append("! colspan=""2"" | ")
                SB.Append(ColumnNames(Columns(0)))
                For Count = 1 To Columns.Count - 1
                    SB.Append("!! ")
                    SB.Append(ColumnNames(Columns(Count)))
                Next
                SW.WriteLine(SB.ToString)

                For Each Item In OutputTitles
                    SB.Clear()
                    SB.AppendLine("|-")
                    SB.Append("| style=""background:rgb(")
                    SB.Append(Item.ToOutput(0))
                    SB.Append(")"" | ")
                    For Each Column In Columns
                        SB.Append("||")
                        SB.Append(Item.ToOutput(Column))
                    Next
                    SW.WriteLine(SB.ToString)
                Next
                SW.WriteLine("|}")
            End Using
        End If
    End Sub
    <Extension()>
    Function ToOutput(Key As Integer, Column As Byte) As String

        Select Case Column
            Case 0
                Return Colours(Key)
            Case 1
                Return Titles(Key) 't
            Case 2
                Return Names(Key) 'n
            Case 3
                Return Lieges(e).StringOrNull(Key) 'le
            Case 4
                Return Lieges(k).StringOrNull(Key) 'lk
            Case 5
                Return Lieges(d).StringOrNull(Key) 'ld
            Case 6
                Return Vassals(k).StringOrNull(Key) 'vk
            Case 7
                Return Vassals(k).CountOrNull(Key) 'cvk
            Case 8
                Return Vassals(d).StringOrNull(Key) 'vd
            Case 9
                Return Vassals(d).CountOrNull(Key) 'cvd
            Case 10
                Return Vassals(c).StringOrNull(Key) 'vc
            Case 11
                Return Vassals(c).CountOrNull(Key) 'cvc
            Case 12
                Return Vassals(b).StringOrNull(Key) 'vb
            Case 13
                Return Vassals(b).CountOrNull(Key) 'cvb
            Case 14
                Return LargestCounty.StringOrNull(Key) 'lc
            Case 15
                Return Capitals(Key) 'cpt
            Case 16
                Return Cultures.StringOrNull(Key) 'clt
            Case 17
                Return Faiths.StringOrNull(Key) 'f
            Case 18
                Return TotalDev.StringOrNull(Key) 'dt
            Case 19
                Return HighestDev.StringOrNull(Key) 'dh
            Case 20
                Return CountyDev.StringOrNull(Key) 'dc
            Case 21
                'sb
                Dim OutputString As String = SpecialBuildings.StringOrNull(Key)
                If Not OutputString.Length = 0 Then
                    Return String.Concat(vbCrLf, "Special Buildings:", OutputString, vbCrLf)
                Else
                    Return ""
                End If
            Case 22
                'm
                Dim OutputString As String = Modifiers.StringOrNull(Key)
                If Not OutputString.Length = 0 Then
                    Return String.Concat(vbCrLf, "Modifiers:", OutputString, vbCrLf)
                Else
                    Return ""
                End If
            Case 23
                'sbm
                Dim SBOutputString As String = SpecialBuildings.StringOrNull(Key)
                Dim MOutputString As String = Modifiers.StringOrNull(Key)
                Dim SB As New StringBuilder
                If Not SBOutputString.Length = 0 Then
                    If Not MOutputString.Length = 0 Then
                        SB.Append(vbCrLf)
                        SB.Append("*"c)
                        SB.Append(" "c)
                        SBOutputString = SBOutputString.Replace("*"c, "**")
                    End If
                    SB.Append($"Special Buildings:{SBOutputString}")
                End If
                If Not MOutputString.Length = 0 Then
                    If Not SBOutputString.Length = 0 Then
                        SB.Append(vbCrLf)
                        SB.Append("*"c)
                        SB.Append(" "c)
                        MOutputString = MOutputString.Replace("*"c, "**")
                    End If
                    SB.Append($"Modifiers:{MOutputString}")
                End If
                If Not SB.Length = 0 Then
                    SB.Append(vbCrLf)
                End If
                Return SB.ToString
        End Select
        Return ""
    End Function
    <Extension()>
    Function StringOrNull(Dict As Dictionary(Of Integer, String), Key As Integer) As String
        If Dict.ContainsKey(Key) Then
            Return Dict(Key)
        Else
            Return ""
        End If
    End Function
    <Extension()>
    Function StringOrNull(Dict As Dictionary(Of Integer, List(Of String)), Key As Integer) As String
        If Dict.ContainsKey(Key) Then
            Return String.Join(", ", Dict(Key))
        Else
            Return ""
        End If
    End Function
    <Extension()>
    Function CountOrNull(Dict As Dictionary(Of Integer, List(Of String)), Key As Integer) As String
        If Dict.ContainsKey(Key) Then
            Return Dict(Key).Count
        Else
            Return ""
        End If
    End Function
    <Extension()>
    Function StringOrNull(Dict As Dictionary(Of Integer, Integer), Key As Integer) As String
        If Dict.ContainsKey(Key) Then
            Return Dict(Key)
        Else
            Return ""
        End If
    End Function
    <Extension()>
    Function StringOrNull(Dict As Dictionary(Of Integer, Dictionary(Of String, Integer)), Key As Integer) As String
        If Dict.ContainsKey(Key) Then
            Dim SB As New StringBuilder
            For Each Item In Dict(Key)
                If Item.Value = 1 Then
                    SB.Append(String.Concat(vbCrLf, "*"c, " "c, Item.Key))
                Else
                    SB.Append(String.Concat(vbCrLf, "*"c, " "c, Item.Key, " "c, "["c, Item.Value, "]"c))
                End If
            Next
            Return SB.ToString
        Else
            Return ""
        End If
    End Function
    Public Function GetNthIndex(searchString As String, charToFind As Char, n As Integer) As Integer
        Dim charIndexPair = searchString.Select(Function(c, i) New With {.Character = c, .Index = i}) _
                                        .Where(Function(x) x.Character = charToFind) _
                                        .ElementAtOrDefault(n - 1)
        Return If(charIndexPair IsNot Nothing, charIndexPair.Index, -1)
    End Function
End Module
