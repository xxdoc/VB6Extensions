Attribute VB_Name = "StringHelperModule"
'Auteur     : Mathieu Guindon
'Création   : 2013-09-18
'Description: Ce module expose de multiples fonctions utilitaires servant à comparer et manipuler des chaînes de caractères (Strings).

Private helper As New StringHelper
Option Explicit

Public Function StringContains(ByVal string_source As String, ByVal find_text As String, Optional ByVal caseSensitive As Boolean = False) As Boolean
'renvoie true si [string_source] contient [find_text]. Fournir true au paramètre optionel [caseSensitive] pour tenir compte de la casse.

    StringContains = helper.StringContains(string_source, find_text, caseSensitive)

End Function

Public Function StringContainsAny(ByVal string_source As String, ByVal caseSensitive As Boolean, ParamArray find_strings()) As Boolean
'renvoie true si [string_source] contient n'importe laquelle des valeurs contenues dans [find_strings()].
    
    Dim valuesArray() As Variant
    valuesArray = find_strings
    StringContainsAny = helper.StringContainsAny(string_source, caseSensitive, valuesArray)

End Function

Public Function StringEndsWith(ByVal string_source As String, ByVal find_text As String, Optional ByVal caseSensitive As Boolean = False) As Boolean
'renvoie true si [string_source] se termine par [find_text]. Fournir true au paramètre optionel [caseSensitive] pour tenir compte de la casse.
    
    StringEndsWith = helper.StringEndsWith(find_text, string_source, caseSensitive)

End Function

Public Function StringEndsWithAny(ByVal string_source As String, ByVal caseSensitive As Boolean, ParamArray find_strings()) As Boolean
'renvoie true si [string_source] se termine par n'importe laquelle des valeurs contenues dans [find_strings()].
    
    Dim valuesArray() As Variant
    valuesArray = find_strings
    StringEndsWithAny = helper.StringEndsWithAny(string_source, caseSensitive, valuesArray)

End Function

Public Function StringFormat(format_string As String, ParamArray values())
'une puissante fonction permettant de formater n'importe quelle valeur en n'importe quel format.
'cette fonction est une implémentation de la fonction [string.Format()] du framework .NET, qui supporte la plupart des fonctionnalités supportées par la version .NET.
'
'exemples: StringFormat("{0}, {1}!", "Hello", "world") => Hello, world!
'          StringFormat("Montant: {0:C2}", 100)        => Montant: $100.00
'          StringFormat("\q{0:N,-10}\q", 10000)        => "10,000    "
'          StringFormat("{0:000000}", 100)             => 000100
'          StringFormat("{0:S}\n{0:D}\n{0:d}", Now)    => 2013-09-18T14:50:49
'                                                         Wednesday, September 18, 2013
'                                                         9/18/2013
'          StringFormat("@C:\SomeFolder\File.txt")     => C:\SomeFolder\File.txt
'          StringFormat("C:\\SomeFolder\\File.txt")    => C:\SomeFolder\File.txt
    '
'spécifier "@" comme premier caractère de [format_string] pour éviter le traitement des séquences d'échappement, ou doubler les backslashes ("\" => "\\") pour simplement échapper un backslash.
'voir [StringHelper].[Class_Initialize] pour les détails sur les formats disponibles et les séquences d'échappement ("\n", "\q", etc.) valides.
    
    Dim valuesArray() As Variant
    valuesArray = values
    StringFormat = helper.StringFormat(format_string, valuesArray)
    
End Function

Public Function StringMatchesAll(ByVal string_source As String, ParamArray find_strings()) As Boolean
'renvoie true si [string_source] correspond exactement à toutes les valeurs contenues dans [find_strings()]. Cette fonction est sensible à la casse.
    
    Dim valuesArray() As Variant
    valuesArray = find_strings
    StringMatchesAll = helper.StringMatchesAll(string_source, valuesArray)

End Function

Public Function StringMatchesAny(ByVal string_source As String, ParamArray find_strings()) As Boolean
'renvoie true si [string_source] correspond exactement à n'importe laquelle des valeurs contenues dans [find_strings()]. Cette fonction est sensible à la casse.
    
    Dim valuesArray() As Variant
    valuesArray = find_strings
    StringMatchesAny = helper.StringMatchesAny(string_source, valuesArray)

End Function

Public Function StringStartsWith(ByVal string_source As String, ByVal find_text As String, Optional ByVal caseSensitive As Boolean = False) As Boolean
'renvoie true si [string_source] débute par [find_text]. Fournir true au paramètre optionel [caseSensitive] pour tenir compte de la casse.
    
    StringStartsWith = helper.StringStartsWith(find_text, string_source, caseSensitive)

End Function

Public Function StringStartsWithAny(ByVal string_source As String, ByVal caseSensitive As Boolean, ParamArray find_strings()) As Boolean
'renvoie true si [string_source] débute par n'importe laquelle des valeurs contenues dans [find_strings()].
    
    Dim valuesArray() As Variant
    valuesArray = find_strings
    StringStartsWithAny = helper.StringStartsWithAny(string_source, caseSensitive, valuesArray)

End Function
