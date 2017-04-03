let
    //FUNKCJE

    //Funkcja Merge(lista) ³¹czy wszystkie rekordy (które s¹ listami) w jedn¹ listê
    //MergeIn to funkcja pomocniczna dla Merge
    MergeIn = (records as list) => List.Generate(
         () => {records,  0, List.Count(records) + 1, {}},
         (_) => _{1} < _{2},
         (_) => {_{0}, _{1} + 1, _{2}, _{3} & _{0}{_{1}}}
    ),
    Merge = (records as list) => MergeIn(records){List.Count(MergeIn(records)) - 1}{3},

    //Funkcja ktora strzela do kolejnych URL z listy i zwraca zawartosc
    LinkToPages = (Link as text) => Json.Document(Web.Contents(Link)),

    //Z ka¿dej strony pobieramu wszystkie watki i laczymy w jedna liste
    PagesToThreads = (BoardName,Pages) => List.Transform(Pages, (_) => 
      List.Transform(_[threads], (_) => Record.AddField(_, "board", BoardName) )
    ),

    //Dany w¹tek przeksz³camy w interesuj¹c¹ nas postaæ, na pocz¹tku trzeba podaæ nazwê boardu, z ktorego jest w¹tek
    ThreadToRecord = (Board as text, InRecord as record) => {
        InRecord[board], 
        InRecord[name],
        InRecord[com], 
        "http://boards.4chan.org/" & InRecord[board] & "/thread/" & Number.ToText(InRecord[no]), 
        "http://i.4cdn.org/" & InRecord[board] & "/" & Number.ToText(InRecord[tim]) & InRecord[ext]
    },



    //MAIN 

    //ród³o (lista boardów)
    Source = Excel.CurrentWorkbook(){[Name="Source"]}[Content],
    SourcesList = Table.ToList(Source),

    //Dostaniemy wszystkie w¹tki (trzeba zmergowac strukture BOARDS -> PAGES -> THREADS
    Threads = Merge(Merge(List.Transform(
        SourcesList,
        (_) => PagesToThreads(_, LinkToPages("http://a.4cdn.org/" & _ & "/catalog.json"))
    ))),
    
    //tera tworzymy ostatecznie tablice wynikowa
    ResultHeader = {"Tablica", "U¿ytkownik", "Komentarz", "W¹tek", "Obrazek"},
    ResultRecords = List.Transform(Threads,(_) => ThreadToRecord("vp", _)),
    ResultTable = #table(ResultHeader, ResultRecords)
in
    ResultTable