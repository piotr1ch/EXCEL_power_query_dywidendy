# EXCEL_power_query_dywidendy

## Opis projektu

Projekt to skoroszyt Excel, który automatyzuje proces pobierania danych z serwisu [strefa inwestorów](https://strefainwestorow.pl/).
Skoroszyt gromadzi informacje o spółkach notowanych na warszawskiej giełdzie (GPW), które od 2000 roku prowadziły politykę dywidendową. Dzięki wykorzystaniu Power Query, arkusz umożliwia szybkie i efektywne pobieranie oraz przetwarzanie tych danych, co pozwala na łatwiejsze śledzenie historii dywidendowej spółek.


## Krok 1: kwerenda z roku 2024

Kwerenda Dywidendy_2024 służy do pobierania i przetwarzania danych o dywidendach z roku 2015. Proces przetwarzania danych obejmuje następujące kroki:

**1. Pobranie zawartości strony.**

**2. Ekstrakcja tabeli HTML.**

**3. Ustawienie nagłówków.**

**4. Zamiana wartości:** Kwerenda dokonuje zamiany znaków w kolumnach "Stopa dywidendy" oraz "Dywidenda na akcję", zamieniając kropki na przecinki i usuwając znaki specjalne (np. "*").

**5. Dostosowanie kolumny "Dzień wypłaty dywidendy":** Komórki w tej kolumnie mogą zawierać do trzech różnych dat. W pierwszym kroku usuwane są wszelkie niepożądane znaki, pozostawiając tylko litery, cyfry oraz kropki. Następnie, wartości w tej kolumnie są dzielone na trzy oddzielne kolumny, każda zawierająca maksymalnie 10 znaków, aby umożliwić łatwiejsze przetwarzanie i analizę danych.

**6. Tworzenie kolumny "Dywidenda [PLN]":** Na podstawie danych w kolumnie "Dywidenda na akcję" tworzona jest nowa kolumna "Dywidenda [PLN]", która wyodrębnia wartość dywidendy w złotych.

**7. Usuwanie kolumny "Dywidenda na akcję":** Po utworzeniu nowej kolumny "Dywidenda [PLN]" usuwana jest oryginalna kolumna "Dywidenda na akcję".

**8. Reorganizacja kolumn:** Na koniec, kolumny są uporządkowane w logiczny sposób, aby ułatwić analizę danych.


```m
let
    // Pobranie zawartości strony internetowej
    Source = Web.BrowserContents("https://strefainwestorow.pl/dane/dywidendy/2015"),

    // Wydobycie danych z tabeli HTML
    #"Extracted Table From Html" = Html.Table(Source, 
        {
            {"Column1", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(1)"},
            {"Column2", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(2)"},
            {"Column3", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(3)"},
            {"Column4", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(4)"},
            {"Column5", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(5)"},
            {"Column6", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(6)"},
            {"Column7", "TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR > :nth-child(7)"}
        },
        [RowSelector="TABLE.table.d-none.d-lg-table.table-dividends-desktop.responsive-enabled.table-hover.table-striped > * > TR"]
    ),

    // Promowanie pierwszego wiersza na nagłówki kolumn
    #"Promoted Headers" = Table.PromoteHeaders(#"Extracted Table From Html", [PromoteAllScalars=true]),

    // Zmiana kropki na przecinek w kolumnach "Stopa dywidendy" i "Dywidenda na akcję"
    #"Replaced Value" = Table.ReplaceValue(#"Promoted Headers", ".", ",", Replacer.ReplaceText, {"Stopa dywidendy", "Dywidenda na akcję"}),

    // Usuwanie znaków "*" w kolumnach "Stopa dywidendy", "Dywidenda na akcję" i "Dzień wypłaty dywidendy"
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value", "*", "", Replacer.ReplaceText, {"Stopa dywidendy", "Dywidenda na akcję", "Dzień wypłaty dywidendy"}),

    // Usuwanie niechcianych znaków w kolumnie "Dzień wypłaty dywidendy" (zostawiamy tylko litery, cyfry i kropki)
    #"Trimmed Text" = Table.TransformColumns(#"Replaced Value", {"Dzień wypłaty dywidendy", each Text.Select(_, {"a".."z", "A".."Z", "0".."9", "."})}),

    // Podział kolumny "Dzień wypłaty dywidendy" na trzy oddzielne kolumny
    #"Split Column by Position" = Table.SplitColumn(#"Trimmed Text", "Dzień wypłaty dywidendy", Splitter.SplitTextByRepeatedLengths(10), 
        {"Dzień wypłaty dywidendy.1", "Dzień wypłaty dywidendy.2", "Dzień wypłaty dywidendy.3"}
    ),

    // Dodanie nowej kolumny "Dywidenda [PLN]" na podstawie danych z "Dywidenda na akcję"
    #"Added Custom" = Table.AddColumn(#"Split Column by Position", "Dywidenda [PLN]", each 
        if Text.Contains([Dywidenda na akcję], "(") then
            Text.BeforeDelimiter(Text.BetweenDelimiters([Dywidenda na akcję], "(", ")"), "zł")
        else 
            Text.Select([Dywidenda na akcję], {"0".."9", ","})
    ),

    // Usunięcie kolumny "Dywidenda na akcję" po obliczeniu wartości
    #"Removed Columns" = Table.RemoveColumns(#"Added Custom", {"Dywidenda na akcję"}),

    // Przekształcenie kolejności kolumn
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns", 
        {"Ticker", "Spółka", "Nazwa", "Dzień dywidendy", "Stopa dywidendy", "Dywidenda [PLN]", "Dzień wypłaty dywidendy.1", "Dzień wypłaty dywidendy.2", "Dzień wypłaty dywidendy.3"}
    )

in
    #"Reordered Columns"

```


W wyniku kwerendy otrzymujemy tabelę z danymi o dywidendach dla spółek za 2024 rok, gotową do dalszej analizy i raportowania.

![Dywidendy_2024](assets/dywidendy_2024.png)


## Krok 2: kwerenda obejmująca okres od roku 2000 do roku poprzedzającego aktualny rok


Aby zebrać dane z kilku stron została zdefiniowana funkcja fxDywidendy, a następnie z jej użyciem utworzona kwerenda z danych z całego wskazanego okresu.


**1. Funkcja fxDywidendy** jest prawie identyczna jak w przypadku wcześniejszej kwerendy, tylko z dodaniem parametru year, który umożliwia dynamiczne wywołanie kwerendy dla dowolnego roku.
 
Kod dla funkcji fxDywidendy:

```m
(year) as table =>

let
    Source = Web.BrowserContents("https://strefainwestorow.pl/dane/dywidendy/" & Text.From(year)),
...(pozostały kod z kwerendy Dywidendy_2024)...
```

**2. Kwerenda:**

- Tworzy listę lat ze wskazanego zakresu, generuje liczby od 2000 do poprzedniego roku (np. w 2025 roku: od 2000 do 2024).

- Dla każdego roku wykonuje następujące operacje:
Konwertuje liczbę na tekst (Text.From(_)) i przypisuje do kolumny "Rok".
- Wywołuje funkcję fxDywidendy(Rok), która zwraca tabelę z danymi dywidend dla danego roku i przypisuje wynik do kolumny "Tabele".
- Konwersja do tabeli: Table.FromRecords(Source) 

W wyniku działania tej kwerendy otrzymana zostaje jedna tabela z danymi dywidend dla wszystkich lat od 2000 do roku poprzedzającego bieżący, gotową do dalszej analizy.

```m
let
    Source = List.Transform({2000..Date.Year(DateTime.LocalNow())-1}, 
        each [ Rok = _, Tabele = fxDywidendy(_) ]  
        // Przekazujemy liczbę do funkcji
    ),
    TableResult = Table.FromRecords(Source),
    //Konwertuje listę rekordów na tabelę
    #"Expanded Tabele" = Table.ExpandTableColumn(TableResult, "Tabele", 
        {"Ticker", "Spółka", "Nazwa", "Dzień dywidendy", "Stopa dywidendy", "Dywidenda [PLN]", "Dzień wypłaty dywidendy.1", "Dzień wypłaty dywidendy.2", "Dzień wypłaty dywidendy.3"}
    ),
    #"Changed Type" = Table.TransformColumnTypes(#"Expanded Tabele",{{"Rok", Int64.Type}, {"Ticker", type text}, {"Spółka", type text}, {"Nazwa", type text}, {"Dzień dywidendy", type date}, {"Stopa dywidendy", Percentage.Type}, {"Dywidenda [PLN]", type number}, {"Dzień wypłaty dywidendy.1", type date}, {"Dzień wypłaty dywidendy.2", type date}, {"Dzień wypłaty dywidendy.3", type date}})
in
    #"Changed Type"
```

![Dywidendy_2024](assets/dywidendy.png)
