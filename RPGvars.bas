Attribute VB_Name = "RPGvars"



Global talked As Boolean

Global gil As Long



Global currentMap(-1 To 20, -1 To 20) As Integer
Global charX As Integer
Global charY As Integer
Global direction As Integer
Global Const NORTH = 0
Global Const EAST = 1
Global Const SOUTH = 2
Global Const WEST = 3
Global mapX, mapY As Integer
Global mapHotSpot(0 To 9, 0 To 9, 0 To 19, 0 To 19) As String
Global hotDoor(0 To 9, 0 To 9, 0 To 19, 0 To 19) As String
Global hotStair(0 To 9, 0 To 9, 0 To 19, 0 To 19, 0 To 1) As Integer
Global keepPos(0 To 9, 0 To 9, 0 To 19, 0 To 19) As Integer

'when using stair has to load a new map
Global loadstair(0 To 9, 0 To 9, 0 To 19, 0 To 19) As String
Global mapName(0 To 9, 0 To 9) As String
Global Const WATER = 0
Global Const BEACH = 1
Global Const LAND = 2
Global Const FOREST = 3
Global Const MOUNTAIN = 4
Global Const BRIDGE = 5
Global Const ROOF = 6
Global Const WALL = 7
Global Const DOOR = 8
Global Const TREEBLOCK = 9
Global Const TOWN = 10
Global Const FLOOR = 11
Global Const FLOOR1 = 12
Global Const STAIR = 13
Global Const PERSON0 = 14

Global mapFile As String

Global charNames(0 To 10) As String
Global townNames(0 To 100) As String

Global speech(0 To 9, 0 To 9, 0 To 19, 0 To 19, 0 To 20) As String
Global progressSpeech(0 To 9, 0 To 9, 0 To 19, 0 To 19, 0 To 20) As Integer
Global affects(0 To 9, 0 To 9, 0 To 19, 0 To 19, 0 To 400) As Integer
Global doEvent(0 To 9, 0 To 9, 0 To 19, 0 To 19, 0 To 20) As Integer





Public Function loadMapInfo()
    For a = 0 To 9
        For b = 0 To 9
            mapName(b, a) = "x" & b & "y" & a & ".txt"
        Next b
    Next a
    mapHotSpot(1, 7, 4, 5) = "myhome.txt"
    hotDoor(1, 7, 13, 11) = "storex13y11.txt"
    hotDoor(1, 7, 4, 16) = "storex4y16.txt"
    hotDoor(1, 7, 11, 4) = "storex11y4.txt"
    hotDoor(1, 7, 3, 7) = "storex3y7.txt"
    
    mapHotSpot(1, 7, 6, 13) = "cabinx1y7.txt"
    hotDoor(1, 7, 14, 14) = "storex14y14.txt"
    hotStair(1, 7, 13, 13, 0) = 3
    hotStair(1, 7, 13, 13, 1) = 3
    hotStair(1, 7, 3, 3, 0) = 13
    hotStair(1, 7, 3, 3, 1) = 13
    
    'HARBOR
    mapHotSpot(2, 8, 1, 8) = "harbortown.txt"
    hotDoor(2, 8, 15, 13) = "harborx15y13.txt"
    hotDoor(2, 8, 8, 4) = "harborx8y4.txt"
    hotDoor(2, 8, 17, 6) = "harborx17y6.txt"
    
    
    'forest lumberJACK
    mapHotSpot(2, 7, 5, 11) = "jackx5y11.txt"
    hotDoor(2, 7, 13, 15) = "jackx13y15.txt"
    hotStair(2, 7, 18, 7, 0) = 4
    hotStair(2, 7, 18, 7, 1) = 4
    hotStair(2, 7, 4, 4, 0) = 18
    hotStair(2, 7, 4, 4, 1) = 7
    
    
    
    'castle stuff
    mapHotSpot(3, 7, 13, 5) = "castlex13y6.txt"
    mapHotSpot(3, 7, 12, 5) = "castlex13y6.txt"
    hotDoor(3, 7, 8, 15) = "castlef0.txt"
    hotDoor(3, 7, 9, 15) = "castlef0.txt"
    hotDoor(3, 7, 10, 15) = "castlef0.txt"
    hotDoor(3, 7, 11, 15) = "castlef0.txt"
    keepPos(3, 7, 8, 15) = 1
    keepPos(3, 7, 9, 15) = 1
    keepPos(3, 7, 10, 15) = 1
    keepPos(3, 7, 11, 15) = 1
    
    loadstair(3, 7, 18, 11) = "castlef1.txt"
    loadstair(3, 7, 9, 18) = "castlef0.txt"
    
    
    
    
'******************************************************************
'*         STORE SECTION         STORE SECTION                    *
'******************************************************************
    storeIDs(1, 7, 15, 16) = 1
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'******************************************************************
'*         DIALOGUE SECTION         DIALOGUE SECTION              *
'******************************************************************
    
    
    ' load dialogue!!
    
    'Store Dialogue
    speech(1, 7, 15, 16, 0) = "Clerk:  Hi.  What can I get you?"
    progressSpeech(1, 7, 15, 16, 0) = 0
    
    
    'mom
    speech(1, 7, 18, 10, 0) = "Mom:  Good morning..."
    progressSpeech(1, 7, 18, 10, 0) = 1
    'GET NAME!
    doEvent(1, 7, 18, 10, 0) = 3
    speech(1, 7, 18, 10, 1) = "Mom:  Please, " & charNames(0) & ", don't go.  You're all I've got left since your father died."
    progressSpeech(1, 7, 18, 10, 1) = 1
    speech(1, 7, 18, 10, 2) = charNames(0) & ":  I've always dreamed of being a soldier and following in dad's footsteps.  I'm going to do this."
    progressSpeech(1, 7, 18, 10, 2) = 1
    speech(1, 7, 18, 10, 3) = "Mom:  ..."
    progressSpeech(1, 7, 18, 10, 3) = 0
    
    
    'soldiers
    speech(1, 7, 8, 18, 0) = "Soldier:  Protecting the village... (Where's the glory in this?)"
    progressSpeech(1, 7, 8, 18, 0) = 0
    
    speech(1, 7, 8, 18, 1) = "Soldier:  They... came.. and killed everyone, even stinky Jill."
    progressSpeech(1, 7, 8, 18, 1) = 1
    
    speech(1, 7, 8, 18, 2) = charNames(0) & ":  Mother..."
    progressSpeech(1, 7, 8, 18, 2) = 1
    
    speech(1, 7, 8, 18, 3) = "Soldier:  I'm sorry, kid...  They're all gone."
    progressSpeech(1, 7, 8, 18, 3) = -1
    
    
    speech(1, 7, 8, 9, 0) = "Soldier:  Wanna be a soldier, huh?  That's a good lad.  We need anyone we can find, now that war is in full swing."
    progressSpeech(1, 7, 8, 9, 0) = 1
    speech(1, 7, 8, 9, 1) = "Soldier:  Might have to wait a while though, kid.  The bridge over the East River got destroyed a short time ago."
    progressSpeech(1, 7, 8, 9, 1) = 1
    speech(1, 7, 8, 9, 2) = "Soldier:  It'll be tough to find a way to the castle.  Good luck."
    progressSpeech(1, 7, 8, 9, 2) = 0
    
    
    'dancing old guys
    speech(1, 7, 11, 14, 0) = "Dancing Pete:  I can't stop dancing, this music's great. Wonder where it's coming from?"
    progressSpeech(1, 7, 11, 14, 0) = 0
    
    speech(1, 7, 8, 13, 0) = "MC Pee Pants:  Someone make him stop... please."
    progressSpeech(1, 7, 8, 13, 0) = 0

    
    
    
    'stinky jill's money (myhome)
    speech(1, 7, 14, 5, 0) = "Stinky Jill:  Mean Jean keeps stealing my money so I hid it in the woods. Shhhh."
    progressSpeech(1, 7, 14, 5, 0) = 0
    
    
    speech(1, 7, 12, 16, 0) = "Mean Jean:  That stinky Jill sure stinks."
    progressSpeech(1, 7, 12, 16, 0) = 0
    
    
    'slutty suzie quest....
    
    speech(1, 7, 14, 15, 0) = "Old Man Stevenson:  I'm too old for this! I need help watching over the fortune in my house. It would be a shame if someone were to come steal it... Please find Suzie for me in town."
    progressSpeech(1, 7, 14, 15, 0) = 0
    
    affects(1, 7, 14, 15, 0) = 1
    affects(1, 7, 14, 15, 1) = 7
    affects(1, 7, 14, 15, 2) = 3
    affects(1, 7, 14, 15, 3) = 10
    

    
    speech(1, 7, 14, 15, 1) = "Old Man Stevenson:  AHH CAPITAL! You've found her. Say, go talk to my brother at the harbor.  He'll reward you for me."
    progressSpeech(1, 7, 14, 15, 1) = 1
    
    'affect dude at harbor
    affects(1, 7, 14, 15, 8) = 2
    affects(1, 7, 14, 15, 9) = 8
    affects(1, 7, 14, 15, 10) = 13
    affects(1, 7, 14, 15, 11) = 14
    
    
    
    speech(1, 7, 14, 15, 2) = "Old Man Stevenson:  Thank you, kind lad."
    progressSpeech(1, 7, 14, 15, 2) = 0
    

    
    speech(1, 7, 14, 15, 3) = "Old Man Stevenson:  Safe and sound..."
    progressSpeech(1, 7, 14, 15, 3) = 0
    
    
    
    speech(1, 7, 3, 10, 0) = "Suzie:  So much to do..."
    progressSpeech(1, 7, 3, 10, 0) = 0
    speech(1, 7, 3, 10, 1) = "Suzie:  Old Man Stevenson, eh?  Tell him I'll be right over."
    progressSpeech(1, 7, 3, 10, 1) = 0
    affects(1, 7, 3, 10, 4) = 1
    affects(1, 7, 3, 10, 5) = 7
    affects(1, 7, 3, 10, 6) = 14
    affects(1, 7, 3, 10, 7) = 15
    speech(1, 7, 3, 10, 2) = "Suzie:  So much to do..."
    progressSpeech(1, 7, 3, 10, 2) = 0
    doEvent(1, 7, 3, 10, 1) = 1
    
    
    'HARBOR TALK!!!!
    'soldier at harbor
    speech(2, 8, 8, 10, 0) = "Soldier:  Gotta keep my eyes peeled.  The enemy is everywhere."
    progressSpeech(2, 8, 8, 10, 0) = 0
    
    
    'find lumber!
    speech(2, 8, 13, 14, 0) = "Older Man Stevenson:  Building boats. Yep, that's all I do... and bridges."
    progressSpeech(2, 8, 13, 14, 0) = 0
    
    speech(2, 8, 13, 14, 1) = "Older Man Stevenson:  My brother?  Yeah, he's kind of old, I guess."
    progressSpeech(2, 8, 13, 14, 1) = 1
    
    speech(2, 8, 13, 14, 2) = "Older Man Stevenson:  Looking for a way across that East River, are ye?  Find me some wood and I'll build a bridge for ye."
    progressSpeech(2, 8, 13, 14, 2) = 0
    
    speech(2, 8, 13, 14, 3) = "Older Man Stevenson:  Why don't you go talk to Timmy at the bar.  He'll know where to get you a good deal on some lumber."
    progressSpeech(2, 8, 13, 14, 3) = 1
    
    'affect bartender
    affects(2, 8, 13, 14, 12) = 2
    affects(2, 8, 13, 14, 13) = 8
    affects(2, 8, 13, 14, 14) = 8
    affects(2, 8, 13, 14, 15) = 16
   
    
    speech(2, 8, 13, 14, 4) = "Older Man Stevenson:  Building boats. Yep, that's all I do... and bridges."
    progressSpeech(2, 8, 13, 14, 4) = 0
    
    '----------------------
    '---at the bar
    
    speech(2, 8, 8, 16, 0) = "Bartender Timmy:  What can I get... Wait a minute! Lemme see some ID, son."
    progressSpeech(2, 8, 8, 16, 0) = 0
    
    speech(2, 8, 8, 16, 1) = "Bartender Timmy:  Go to the forest by the East River.  I know a lumberjack who lives there.  Tell him I sent you."
    progressSpeech(2, 8, 8, 16, 1) = 0
        
    'affect lumberjack!!!
    affects(2, 8, 8, 16, 4) = 2
    affects(2, 8, 8, 16, 5) = 7
    affects(2, 8, 8, 16, 6) = 15
    affects(2, 8, 8, 16, 7) = 8
    
    
    
    speech(2, 8, 8, 16, 2) = "Bartender Timmy:  No ID, no booze."
    progressSpeech(2, 8, 8, 16, 2) = 0

    

    speech(2, 8, 12, 13, 0) = "Drunken Bill:  <BELCH>  I'm drunk enough to start brawling soon... with Carl again."
    progressSpeech(2, 8, 12, 13, 0) = 0
    
    speech(2, 8, 16, 13, 0) = "Carl:  Someone's gonna have to shut that Bill up."
    progressSpeech(2, 8, 16, 13, 0) = 0
    
    
    
    
    speech(2, 8, 5, 5, 0) = "Stompy Jones:  Gonna stomp these flowers good."
    progressSpeech(2, 8, 5, 5, 0) = 1
    
    speech(2, 8, 5, 5, 1) = "Stompy Jones:  I'll do what I want, jerk."
    progressSpeech(2, 8, 5, 5, 1) = 0
    
    speech(2, 8, 16, 17, 0) = "Willy:  It's friggin' fridged in here."
    progressSpeech(2, 8, 16, 17, 0) = 0
    
    
    
    'Lumber JACKS HOUSE
    
    speech(2, 7, 15, 8, 0) = "Lumberjack Jack: I put on my flannel today.  What are you doing in my house?"
    progressSpeech(2, 7, 15, 8, 0) = 0
    
    speech(2, 7, 15, 8, 1) = "Lumberjack Jack: Oh Timmy! We go way back.  I owe him a few, I do.  Sure I'll fix you up a bridge right away."
    progressSpeech(2, 7, 15, 8, 1) = 1
    doEvent(2, 7, 15, 8, 1) = 2
    
    
    speech(2, 7, 15, 8, 2) = "Lumberjack Jack: Am I a great lumberjack or what!?"
    progressSpeech(2, 7, 15, 8, 2) = -1


'***************************************************************
'*   CASTLE        CASTLE    FIRST CASTLE                      *
'***************************************************************
    speech(3, 7, 7, 14, 0) = "Soldier: We've failed..."
    speech(3, 7, 11, 14, 0) = "Soldier: Everything's gone..."
    speech(3, 7, 7, 10, 0) = "Soldier: ..."
    speech(3, 7, 11, 10, 0) = "Soldier: Shit."
    speech(3, 7, 7, 6, 0) = "Soldier: ..."
    speech(3, 7, 11, 6, 0) = "Soldier: ..."

    'chancellor
    'events to make------
    'ruin myhome
    
    
    
    speech(3, 7, 9, 4, 0) = "Chancellor:  Who are you?  You must leave at once!"
    progressSpeech(3, 7, 9, 4, 0) = 1
    
    doEvent(3, 7, 9, 4, 1) = 4
    'affects soldier at home
    affects(3, 7, 9, 4, 0) = 1
    affects(3, 7, 9, 4, 1) = 7
    affects(3, 7, 9, 4, 2) = 8
    affects(3, 7, 9, 4, 3) = 18
    
    'affect mayor of next town to go to
    '
    '
    '
    '
    
    speech(3, 7, 9, 4, 1) = charNames(0) & ":  I... I want to become a soldier.  What happened here."
    progressSpeech(3, 7, 9, 4, 1) = 1
    
    affects(3, 7, 9, 4, 4) = 3
    affects(3, 7, 9, 4, 5) = 7
    affects(3, 7, 9, 4, 6) = 8
    affects(3, 7, 9, 4, 7) = 4
   
    speech(3, 7, 9, 4, 2) = "Chancellor:  They've ransacked the castle, taken the king and Princess.  Our kingdom is in ruin."
    progressSpeech(3, 7, 9, 4, 2) = 1
    
    speech(3, 7, 9, 4, 3) = "Chancellor:  We're all that's left."
    progressSpeech(3, 7, 9, 4, 3) = 1
   
    speech(3, 7, 9, 4, 4) = "Chancellor:  You wish to be a soldier, do you...  I will send you on a mission then.  Go warn the people of " & townNames(5) & " of the impending attacks."
    progressSpeech(3, 7, 9, 4, 4) = 1
    
    speech(3, 7, 9, 4, 5) = "Chancellor: My general will go with you."
    progressSpeech(3, 7, 8, 4, 5) = 1
    
    speech(3, 7, 9, 4, 6) = charNames(0) & ": Anything you say, sir."
    progressSpeech(3, 7, 9, 4, 6) = 0


    'get a new character!!
    speech(3, 7, 8, 4, 0) = "????: ..."
    progressSpeech(3, 7, 8, 4, 0) = 0
    
    speech(3, 7, 8, 4, 1) = "????: Let's go.  Why don't we go warn your hometown first."
    progressSpeech(3, 7, 8, 4, 1) = 1
    
    doEvent(3, 7, 8, 4, 1) = 5
    
    speech(3, 7, 8, 4, 2) = "There's no time to waste."
    progressSpeech(3, 7, 8, 4, 2) = -1
    
    
    


  
End Function
