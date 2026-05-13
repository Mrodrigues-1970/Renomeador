Option Explicit
Dim Str_versao
'Renomeador de arquivos 
'versão 1.8 - 30/12/2010
'versão 1.9 - 21/02/2011 - Método Contagem, About
'versão 2.0 - 04/04/2011 - Opção Renomeação numérica de diretórios
'versão 2.1 - 04/04/2011 - Opções de Conversão de Espaços e Underlines
'versão 2.2 - 23/07/2011 - Opção Subtração no nome de diretórios, report considera tipo de objeto
'versão 2.3 - 09/03/2012 - Opção Replace de substring no nome
'versão 2.4 - 19/03/2012 - Verificação existência de nome no modo 12
'versão 2.5 - 20/03/2012 - Manutenção no nome das variaveis
'versão 2.6 - 27/03/2012 - Reestruturação do Help
'versão 2.7 - 29/04/2012 - Correção trailing zeros da opção 02
'versão 2.8 - 03/05/2012 - Opção Alteração de caixa (case)
'versão 2.9 - 25/06/2012 - Verificação de numérico permite números tipo Double
'versão 3.0 - 24/09/2012 - Opção Limpeza de identificadoes Tumblr
'versão 3.1 - 07/07/2013 - Opção Inclusão de prefixo
'versão 3.2 - 09/02/2014 - Opção Retira acentuação
'versão 3.3 - 07/03/2014 - Correção na mensagem de quantidade de alterações
'versão 3.4 - 28/07/2014 - Finaliza Rename Multiplo de arquivos do Windows Explorer
'versão 3.5 - 16/07/2015 - Finaliza Rename Multiplo de diretórios do Windows Explorer
'versão 3.6 - 16/07/2015 - Renomeação parcial de diretórios
'versão 3.7 - 06/06/2016 - Adição no dicionario de palavras descartáveis do Tumblr
'versão 3.8 - 24/07/2018 - Opção Left Trim
'versão 3.9 - 24/04/2020 - Correção bug na opção 15 (loop infinito na coleção de arquivos)
'versão 4.0 - 23/08/2022 - Correção bug no loop da opção 2 (mais de 1 loop no diretório)
Str_versao = "versão 4.0 - 23 agosto 2022"

Dim Arr_NomesArquivos
Dim Arr_PartesNome
Dim Arr_Palavras
Dim iPalavra
Dim Bln_Pendente
Dim Bln_Reincidencia
Dim Int_Afetados
Dim Int_Digitos
Dim Int_Indice
Dim Int_JaExistente
Dim Int_Numero
Dim Int_QtdeDigitos
Dim Int_QtdArquivosFolder
Dim Int_QtdArquivosLimiteAviso
Dim Obj_Arquivo
Dim Obj_Arquivos
Dim Obj_Diretorio
Dim Obj_Diretorios
Dim Obj_EsteFolder
Dim Obj_FileSystem
Dim Str_AntigoNome
Dim Str_Auxiliar
Dim Str_Enter
Dim Str_Extencao
Dim Str_ExtencaoAnterior
Dim Str_ExtencaoSubstituta
Dim Str_ListaNomes
Dim Str_Mensagem
Dim Str_Metodo
Dim Str_MenorNomeCorrente
Dim Str_MeuPath
Dim Str_NomeDesteScript
Dim Str_NomeFinal
Dim Str_NomeSubstituto
Dim Str_NovoNome
Dim Str_Opcao
Dim Str_Palavra
Dim Str_PalavraDestino
Dim Str_PalavrasTumblr
Dim Str_PalavraOrigem
Dim Str_PathCompleto
Dim Str_Resposta
Dim Str_TipoAlteracaoCaixa
Dim Str_TipoObjeto
Dim Str_ZerosEsquerda

Str_Enter = chr(13)
'Quantidade chave para avisos
Int_QtdArquivosLimiteAviso = 30
Bln_Pendente = True
'Não sai enquanto não for escolhida uma opção válida ou Esc
While Bln_Pendente
    Call MensagemMenu
    Str_Opcao = Trim(InputBox(Str_Mensagem,"Renomeador de arquivos"))
    Select case LCase(Str_Opcao)
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"
            Bln_Pendente = False
        Case "?"
            AlertaInfo("Marcelo Rodrigues" & Str_Enter &  "Copyright 2010 - " & DatePart("YYYY", Now) & Str_Enter & "Para ajuda, use ? + o número da opção.")
        Case "?1", "?2", "?3", "?4", "?5", "?6", "?7", "?8", "?9", "?10", "?11", "?12", "?13", "?14", "?15", "?16", "?17", "?18", "?19", "?20"
            Call Ajuda(Str_opcao)
        Case ""
            'Fim do script
            WScript.Quit
        Case "ver", "versao", "versão", "version", "??", "about", "sobre"
            AlertaInfo("Marcelo Rodrigues" & Str_Enter &  "Copyright 2010 - " & DatePart("YYYY", Now) & Str_Enter & Str_versao)
        Case Else
            Call msgbox("Opção '" & Str_Opcao & "' não é válida.", VBCritical,"Opção inválida.")
    End Select
WEnd
    
Str_MeuPath = WScript.ScriptFullName
Int_Indice = InStrRev(Str_MeuPath,"\")
Str_NomeDesteScript = Mid(Str_MeuPath,Int_Indice + 1)
Str_MeuPath = Mid(Str_MeuPath,1,Int_Indice)
Set Obj_FileSystem = CreateObject("Scripting.FileSystemObject")
Set Obj_EsteFolder = Obj_FileSystem.GetFolder(Str_MeuPath)
Set Obj_Arquivos = Obj_EsteFolder.Files
Int_QtdArquivosFolder = Obj_Arquivos.Count
Int_Afetados = 0
Int_JaExistente = 0
Bln_Reincidencia = False

'===== executa opção escolhida ======
Select case Str_Opcao
    Case "1"
        Str_Metodo = "Zero à esquerda"
        Str_TipoObjeto = "arquivos"
        Str_Palavra = PerguntaPalavra("Qual a parte não numérica do nome do arquivo?", True)
        For Each Obj_Arquivo in Obj_Arquivos
            Str_AntigoNome = Obj_Arquivo.name
            'Retira a extenção
            Str_Extencao = RetornaStr_Extencao(Str_AntigoNome)
            Str_Auxiliar = Replace(Str_AntigoNome,Str_Extencao,"")
            'Retira a palavra fornecida
            Str_Auxiliar = Replace(Str_Auxiliar,Str_Palavra,"")
            'Se retirando a palavra e a extenção, sobra um número
            If IsNumeric(Str_Auxiliar) then
                'Se esse número for menor que 10
                If CDbl(Str_Auxiliar) < 10 then
                    Str_NomeSubstituto = Str_Palavra & "0" & Str_Auxiliar & Str_Extencao
                    'Se o novo nome não existir
                    If Not ArquivoExiste(Str_NomeSubstituto) Then
                        Obj_Arquivo.name = Str_NomeSubstituto
                        Int_Afetados = Int_Afetados + 1
                    Else
                        Int_JaExistente = Int_JaExistente + 1
                    End If
                End If
            End If
        Next 

        
    Case "2" 'ordem alfabética necessária
        Str_Metodo = "Substituição total"
        Str_TipoObjeto = "arquivos"
        If Int_QtdArquivosFolder > Int_QtdArquivosLimiteAviso Then
            Str_Mensagem = "Todos os " & (Obj_Arquivos.Count - 1)  & " arquivos serão renomeados. Tem certeza?"
            Str_Resposta = MsgBox(Str_Mensagem, VBYesNo, "Alteração em TODOS os arquivos")
            If Str_Resposta <> VBYes Then
                WScript.Quit
            End If
        End If
        'Quantidades de dígitos da qtde de arquivos (exceto ESTE SCRIPT)
        Int_QtdeDigitos = Len(Int_QtdArquivosFolder - 1)
        Str_ZerosEsquerda = String(Int_QtdeDigitos,"0")

        Str_NomeSubstituto = InputBox("Qual o novo prefixo de nome que TODOS os arquivos receberão?", Str_Metodo)
        If Trim(Str_NomeSubstituto) = "" then
            Call msgbox("Nenhum nome fornecido", VBCritical)
            WScript.Quit
        End If
        Int_Numero = 0
		
		Call MontaListaArquivos
		For Each Str_PathCompleto in Arr_NomesArquivos
			Set Obj_Arquivo = Obj_FileSystem.GetFile(Str_PathCompleto)
			If Obj_Arquivo.name <> Str_NomeDesteScript Then
				Str_AntigoNome = Obj_Arquivo.name
                Str_Extencao = RetornaStr_Extencao(Str_AntigoNome)
                Int_Numero = Int_Numero + 1
				Str_NovoNome = Str_NomeSubstituto & Right(Str_ZerosEsquerda & Int_Numero,Int_QtdeDigitos) & Str_Extencao
				
                If Not ArquivoExiste(Str_NovoNome) Then
                    Obj_Arquivo.name = Str_NovoNome
                    Int_Afetados = Int_Afetados + 1
                Else
                    Int_JaExistente = Int_JaExistente + 1
                End If
			End If
		Next
		
		

    Case "3" 'ordem alfabética necessária
        Str_Metodo = "Substituição parcial"
        Str_TipoObjeto = "arquivos"
        Str_Palavra = PerguntaPalavra("Que palavra os arquivos desejados possuem no nome?", False)
        Str_NomeSubstituto = InputBox("Substituir pelo o que?", Str_Metodo)
        If Trim(Str_NomeSubstituto) = "" then
            Call msgbox("Nenhum nome fornecido", VBCritical)
            WScript.Quit
        End If
        Int_Numero = 0
        For Each Obj_Arquivo in Obj_Arquivos
            If Obj_Arquivo.name <> Str_NomeDesteScript Then
                Str_AntigoNome = Obj_Arquivo.name
                'Se o nome do arquivo possuir a palavra fornecida ENTÃO efetua troca de nome
                If InStr(Str_AntigoNome, Str_Palavra) > 0 Then
                    Str_Extencao = RetornaStr_Extencao(Str_AntigoNome)
                    Int_Numero = Int_Numero + 1
                    If Int_Numero < 10 then
                        Str_NovoNome = Str_NomeSubstituto & "0" & Int_Numero & Str_Extencao
                    else
                        Str_NovoNome = Str_NomeSubstituto & Int_Numero & Str_Extencao
                    End If
                    If Not ArquivoExiste(Str_NovoNome) Then
                        Obj_Arquivo.name = Str_NovoNome
                        Int_Afetados = Int_Afetados + 1
                    Else
                        Int_JaExistente = Int_JaExistente + 1
                    End If
                End If
            End If
        Next


    Case "4"
        Str_Metodo = "Subtração"
        Str_TipoObjeto = "arquivos"
        Str_Palavra = PerguntaPalavra("O que deve ser subtraído?", False)
        For Each Obj_Arquivo in Obj_Arquivos
            Str_Extencao = RetornaStr_Extencao(Obj_Arquivo.name)
            Str_AntigoNome = Replace(Obj_Arquivo.name,Str_Extencao,"")
            If InStr(Str_AntigoNome, Str_Palavra) > 0 Then
                Str_NomeSubstituto = Replace(Str_AntigoNome,Str_Palavra,"")
                If Obj_Arquivo.name <> Str_NomeDesteScript and Obj_Arquivo.name <> Str_NomeSubstituto Then
                    If Not ArquivoExiste(Str_NomeSubstituto & Str_Extencao) Then
                        Obj_Arquivo.name = Str_NomeSubstituto & Str_Extencao
                        Int_Afetados = Int_Afetados + 1
                    Else
                        Int_JaExistente = Int_JaExistente + 1
                    End If
                End If
            End If
        Next 


    Case "5"
        Str_Metodo = "Simplificação de Extenção"
        Str_TipoObjeto = "arquivos"
        For Each Obj_Arquivo in Obj_Arquivos
            Str_Extencao = RetornaStr_Extencao(Obj_Arquivo.name)
            Str_AntigoNome = Replace(Obj_Arquivo.name,Str_Extencao,"")
            Str_Extencao = UCase(Str_Extencao)
            Str_NomeSubstituto = ""
            Select Case Str_Extencao
                Case ".JPEG"
                    Str_NomeSubstituto = Str_AntigoNome & ".jpg"
                Case ".MPEG"
                    Str_NomeSubstituto = Str_AntigoNome & ".mpg"
                Case ".HTML"
                    Str_NomeSubstituto = Str_AntigoNome & ".htm"
                Case Else
                    Str_NomeSubstituto = Obj_Arquivo.name
            End Select
            If Obj_Arquivo.name <> Str_NomeDesteScript and Obj_Arquivo.name <> Str_NomeSubstituto Then
                If Not ArquivoExiste(Str_NomeSubstituto) Then
                    Obj_Arquivo.name = Str_NomeSubstituto
                    Int_Afetados = Int_Afetados + 1
                Else
                    Int_JaExistente = Int_JaExistente + 1
                End If
            End If
        Next

    Case "6"
        Str_Metodo = "Mudança de Extenção"
        Str_TipoObjeto = "arquivos"
        Str_ExtencaoAnterior = PerguntaPalavra("Qual a extenção existente?", False)
        Str_ExtencaoAnterior = "." & Str_ExtencaoAnterior
        Str_ExtencaoAnterior = Replace(Str_ExtencaoAnterior,"..",".")

        Str_ExtencaoSubstituta = PerguntaPalavra("Qual a extenção substituta?", False)
        Str_ExtencaoSubstituta = Replace(Str_ExtencaoSubstituta,".","")

        For Each Obj_Arquivo in Obj_Arquivos
            Str_Extencao = RetornaStr_Extencao(Obj_Arquivo.name)
            Str_AntigoNome = Replace(Obj_Arquivo.name,Str_Extencao,"")
            If Ucase(Str_ExtencaoAnterior) = Ucase(Str_Extencao) Then
                Obj_Arquivo.name = Str_AntigoNome & "." & Str_ExtencaoSubstituta
                Int_Afetados = Int_Afetados + 1
            End If
        Next

    Case "7"
        Str_Metodo = "Contagem"
        Str_TipoObjeto = "arquivos"
        Str_Palavra = PerguntaPalavra("Qual o subnome desejado?", False)
        For Each Obj_Arquivo in Obj_Arquivos
            If InStr(UCase(Obj_Arquivo.name), UCase(Str_Palavra)) > 0 Then
                Int_Afetados = Int_Afetados + 1
            End If
        Next
    
    Case "8"
        Str_Metodo = "Renumeração de diretórios"
        Str_TipoObjeto = "diretórios"
        Int_Numero = PerguntaPalavra("Qual o número inicial?", False)
        If Not IsNumeric(Int_Numero) then
                Call msgbox("Número inválido", VBCritical)
                WScript.Quit
        End if
        Set Obj_Diretorios = Obj_EsteFolder.Subfolders
        Int_Digitos = Len(Obj_Diretorios.Count)
        Str_ZerosEsquerda = Replace(Space(Int_Digitos - 1)," ","0")
        For Each Obj_Diretorio in Obj_Diretorios
            Obj_Diretorio.Name = Left(Str_ZerosEsquerda,Int_Digitos - Len(Int_Numero)) & Int_Numero
            Int_Numero = Int_Numero + 1
            Int_Afetados = Int_Afetados + 1
        Next

    Case "9"
        Str_Metodo = "Espaços para Underline"
        Str_TipoObjeto = "arquivos"
        For Each Obj_Arquivo in Obj_Arquivos
            If InStr(Obj_Arquivo.name, Space(1)) > 0 Then
                Obj_Arquivo.name = Replace(Obj_Arquivo.name, space(1), "_")
                Int_Afetados = Int_Afetados + 1
            End if
        Next
        
    Case "10"
        Str_Metodo = "Underline para Espaços"
        Str_TipoObjeto = "arquivos"
        For Each Obj_Arquivo in Obj_Arquivos
            If InStr(Obj_Arquivo.name, "_") > 0 Then
                Obj_Arquivo.name = Replace(Obj_Arquivo.name, "_", space(1))
                Int_Afetados = Int_Afetados + 1
            End if
        Next

    Case "11"
        Str_Metodo = "Subtração no nome de diretórios"
        Str_TipoObjeto = "diretórios"
        Str_Palavra = PerguntaPalavra("O que deve ser subtraído?", False)
        Set Obj_Diretorios = Obj_EsteFolder.Subfolders
        For Each Obj_Diretorio in Obj_Diretorios
            If InStr(Obj_Diretorio.name, Str_Palavra) > 0 Then
                Str_NovoNome = Replace(Obj_Diretorio.name, Str_Palavra, "")
                If Len(Str_NovoNome) > 0 then
                    Obj_Diretorio.name = Str_NovoNome
                    Int_Afetados = Int_Afetados + 1
                End If
            End If
        Next
    
    Case "12"
        Str_Metodo = "Replace"
        Str_TipoObjeto = "arquivos"
        Str_Palavra = PerguntaPalavra("Que substring os arquivos desejados possuem?" & Str_Enter & "(* para todos)", False)
        Str_PalavraOrigem = PerguntaPalavra("Que substring deseja substituir?", False)
        Str_PalavraDestino = PerguntaPalavra("Qual a nova substring que vai substituir '" & Str_PalavraOrigem & "'?", False)
        For Each Obj_Arquivo in Obj_Arquivos
            If InStr(UCase(Obj_Arquivo.name), UCase(Str_Palavra)) > 0  and (Str_NomeDesteScript <> Obj_Arquivo.name) Then
                Str_NovoNome = Replace(Obj_Arquivo.name, Str_PalavraOrigem, Str_PalavraDestino)
                If ArquivoExiste(Str_NovoNome) Then
                    Call msgbox("O nome '" & Str_NovoNome & "' já existe.", VBCritical)
                    WScript.Quit
                Else
                    Obj_Arquivo.name = Str_NovoNome
                    Int_Afetados = Int_Afetados + 1
                End If
            End If
        Next
    
    Case "13"
        Str_Metodo = "Alteração de caixa"
        Str_TipoObjeto = "arquivos"
        Str_TipoAlteracaoCaixa = PerguntaPalavra("Qual o tipo de alteração?" & Str_Enter & "(A)lta, (B)aixa ou (C)apitalizar", false)
        Str_TipoAlteracaoCaixa = UCase(Str_TipoAlteracaoCaixa)
        If Str_TipoAlteracaoCaixa <> "A" And Str_TipoAlteracaoCaixa <> "B" And Str_TipoAlteracaoCaixa <> "C" Then
            Call msgbox("Opção de alteração de caixa inválida.", VBCritical)
            WScript.Quit
        End If

        For Each Obj_Arquivo in Obj_Arquivos
            If Str_NomeDesteScript <> Obj_Arquivo.name Then
                Select Case Str_TipoAlteracaoCaixa
                    Case "A"
                        Str_NovoNome = UCase(Obj_Arquivo.name)
                    Case "B"
                        Str_NovoNome = LCase(Obj_Arquivo.name)
                    Case "C"
                        Str_NovoNome = CapitalizaNome(Obj_Arquivo.name)
                End Select

                Obj_Arquivo.name = "temp" & Obj_Arquivo.name
                Obj_Arquivo.name = Str_NovoNome
                Int_Afetados = Int_Afetados + 1

            End If
        Next
    
    Case "14"
        Str_Metodo = "Tumblr Clean"
        Str_TipoObjeto = "arquivos"
		Bln_Reincidencia = True
        Str_PalavrasTumblr = "tumblr_,qcru98o1,_1280,_500,_400,_250,_540,_640"
        Arr_Palavras = Split(Str_PalavrasTumblr, ",")
        
        For Each iPalavra in Arr_Palavras
            For Each Obj_Arquivo in Obj_Arquivos
                Str_Extencao = RetornaStr_Extencao(Obj_Arquivo.name)
                Str_AntigoNome = Replace(Obj_Arquivo.name,Str_Extencao,"")
                If InStr(Str_AntigoNome, iPalavra) > 0 Then
                    Str_NomeSubstituto = Replace(Str_AntigoNome,iPalavra,"")
                    If Obj_Arquivo.name <> Str_NomeDesteScript and Obj_Arquivo.name <> Str_NomeSubstituto Then
                        If Not ArquivoExiste(Str_NomeSubstituto & Str_Extencao) Then
                            Obj_Arquivo.name = Str_NomeSubstituto & Str_Extencao
                            Int_Afetados = Int_Afetados + 1
                        Else
                            Int_JaExistente = Int_JaExistente + 1
                        End If
                    End If
                End If
            Next
            Set Obj_Arquivo = Nothing
        Next

    Case "15"
        Str_Metodo = "Inclusão de prefixo"
        Str_TipoObjeto = "arquivos"
		Str_Palavra = PerguntaPalavra("Qual o prefixo desejado?", False)
		Call MontaListaArquivos
		For Each Str_PathCompleto in Arr_NomesArquivos
			Set Obj_Arquivo = Obj_FileSystem.GetFile(Str_PathCompleto)
			Obj_Arquivo.name = Str_Palavra & Obj_Arquivo.name
			Int_Afetados = Int_Afetados + 1
		Next
		

    Case "16"
        Str_Metodo = "Retira Acentuação"
        Str_TipoObjeto = "arquivos"
		Bln_Reincidencia = True
        For Each Obj_Arquivo in Obj_Arquivos
			If Obj_Arquivo.name <> Str_NomeDesteScript Then
                Str_NovoNome = RetiraAcento(Obj_Arquivo.name)
                If Str_NovoNome <> Obj_Arquivo.name Then
                    Obj_Arquivo.name = Str_NovoNome
                    Int_Afetados = Int_Afetados + 1
                End If
			End IF
        Next
        
    Case "17"
        Str_Metodo = "Rename Múltiplo W.Explorer"
        Str_TipoObjeto = "arquivos"
		Bln_Reincidencia = False
        For Each Obj_Arquivo in Obj_Arquivos
            'Se arquivo não for este script E tiver um parêntese no nome
			If (Obj_Arquivo.name <> Str_NomeDesteScript) And (InStr(Obj_Arquivo.name, ")") > 0) Then
                Str_NovoNome = Replace(Obj_Arquivo.name, ")", "")
                Str_NovoNome = Replace(Str_NovoNome, " (", "")
                Obj_Arquivo.name = Str_NovoNome
                Int_Afetados = Int_Afetados + 1
			End IF
        Next

    Case "18"
        Str_Metodo = "Rename Múltiplo W.Explorer Diretórios"
        Str_TipoObjeto = "diretórios"
		Bln_Reincidencia = False
		Set Obj_Diretorios = Obj_EsteFolder.Subfolders
		Int_QtdArquivosFolder = Obj_Diretorios.Count
        For Each Obj_Diretorio in Obj_Diretorios
			If InStrRev(Obj_Diretorio.name, "(") > 0 Then
				Str_NovoNome = Replace(Obj_Diretorio.name, ")", "")
				Str_NovoNome = Replace(Str_NovoNome, " (", "")
				Obj_Diretorio.name = Str_NovoNome
				Int_Afetados = Int_Afetados + 1
			End If
        Next


    Case "19" 'ordem alfabética necessária
        Str_Metodo = "Substituição parcial diretórios"
        Str_TipoObjeto = "diretórios"
        Str_Palavra = PerguntaPalavra("Que palavra os folders desejados possuem no nome?", False)
        Str_NomeSubstituto = InputBox("Substituir pelo o que?", Str_Metodo)
        If Trim(Str_NomeSubstituto) = "" then
            Call msgbox("Nenhum nome fornecido", VBCritical)
            WScript.Quit
        End If
        Int_Numero = 0
		Set Obj_Diretorios = Obj_EsteFolder.Subfolders
        For Each Obj_Diretorio in Obj_Diretorios
			Str_AntigoNome = Obj_Diretorio.name
			'Se o nome do folder possuir a palavra fornecida ENTÃO efetua troca de nome
			If InStr(Str_AntigoNome, Str_Palavra) > 0 Then
				Int_Numero = Int_Numero + 1
				If Int_Numero < 10 then
					Str_NovoNome = Str_NomeSubstituto & "0" & Int_Numero
				else
					Str_NovoNome = Str_NomeSubstituto & Int_Numero
				End If
				If Not FolderExiste(Str_NovoNome) Then
					Obj_Diretorio.name = Str_NovoNome
					Int_Afetados = Int_Afetados + 1
				Else
					Int_JaExistente = Int_JaExistente + 1
				End If
			End If
        Next

    Case "20"
        Str_Metodo = "Left Trim"
        Str_TipoObjeto = "arquivos"
        Str_Palavra = PerguntaPalavra("Qual o novo início?", False)
        For Each Obj_Arquivo in Obj_Arquivos
            Str_Extencao = RetornaStr_Extencao(Obj_Arquivo.name)
            Str_AntigoNome = Replace(Obj_Arquivo.name,Str_Extencao,"")
            If InStr(Str_AntigoNome, Str_Palavra) > 0 Then
                Int_Indice = InStr(Str_AntigoNome, Str_Palavra)
				Str_NomeSubstituto = Mid(Str_AntigoNome,Int_Indice)
                If Obj_Arquivo.name <> Str_NomeDesteScript and Obj_Arquivo.name <> Str_NomeSubstituto Then
                    If Not ArquivoExiste(Str_NomeSubstituto & Str_Extencao) Then
                        Obj_Arquivo.name = Str_NomeSubstituto & Str_Extencao
                        Int_Afetados = Int_Afetados + 1
                    Else
                        Int_JaExistente = Int_JaExistente + 1
                    End If
                End If
            End If
        Next
		
End Select

Set Obj_FileSystem = Nothing
Str_Mensagem = "Método: " & Str_Metodo & Str_Enter & Str_Enter
If Bln_Reincidencia Then
	Str_Mensagem = Str_Mensagem & "Quantidade de alterações nos " & Str_TipoObjeto & ": " & Int_Afetados & Str_Enter
Else
	Str_Mensagem = Str_Mensagem & "Quantidade de " & Str_TipoObjeto & " afetados: " & Int_Afetados & Str_Enter
	Str_Mensagem = Str_Mensagem & "Quantidade de " & Str_TipoObjeto & " já existentes: " & Int_JaExistente & Str_Enter
End If
Str_Mensagem = Str_Mensagem & "Quantidade total de " & Str_TipoObjeto & " analizados: " & Int_QtdArquivosFolder
Call msgbox(Str_Mensagem, VBExclamation, "Relatório do Renomeador")

'Termina o Script
WScript.Quit

'============== Funções ===============================
'Pergunta a palavra principal para ser usada no método
Private Function PerguntaPalavra(Pergunta, blnSugestao)
    Dim lPalavra
    If blnSugestao Then
        lPalavra = InputBox(Pergunta, Str_Metodo, SugestaoNome)
    Else
        lPalavra = InputBox(Pergunta, Str_Metodo)
    End If
    If Trim(lPalavra) = "" then
        Call msgbox("Nenhuma palavra fornecida", VBCritical)
        WScript.Quit
    End If
    PerguntaPalavra = lPalavra
End Function

'Retorna a extenção do arquivo
Private Function RetornaStr_Extencao(NomeArquivo)
    Str_Extencao = ""
    Int_Indice = InStrRev(NomeArquivo,".")
    If Int_Indice > 0 then
        Str_Extencao = Mid(NomeArquivo,Int_Indice)
    End If
    RetornaStr_Extencao = Str_Extencao
End Function

'TENTA sugerir a palavra default pegando o nome do primeiro arquivo (exceto este script)
Private Function SugestaoNome
    For Each Obj_Arquivo in Obj_Arquivos
        If Obj_Arquivo.name <> Str_NomeDesteScript Then
            SugestaoNome = Replace(Obj_Arquivo.name, RetornaStr_Extencao(Obj_Arquivo.name), "")
            Exit For
        End If
    Next
End Function

'Mostra a ajuda na Tela
Private Sub Ajuda(Str_opcao)

    Select Case Str_opcao
    
        Case "?1"
            Str_Mensagem = "1 = ZERO À ESQUERDA" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Corrige a parte numérica do nome dos arquivos com numeração " & Str_Enter 
            Str_Mensagem = Str_Mensagem & "inferior à 10 que não possuem o zero à esquerda." & Str_Enter 
            
        Case "?2"
            Str_Mensagem = "2 = SUBSTITUIÇÃO TOTAL" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Substitui o nome de TODOS os arquivos por uma palavra + contador." & Str_Enter 
    
        Case "?3"
            Str_Mensagem = "3 = SUBSTITUIÇÃO PARCIAL" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "1)Através de uma substring seleciona os arquivos que serão afetados." & Str_Enter 
            Str_Mensagem = Str_Mensagem & "2)Pergunta a nova substring que substuirá a nome do arquivo." & Str_Enter 

        Case "?4"
            Str_Mensagem = "4 = SUBTRAÇÃO" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Retira uma palavra do nome do arquivo." & Str_Enter 

        Case "?5"
            Str_Mensagem = "5 = SIMPLIFICAÇÃO DE EXTENÇÃO" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Converte:" & Str_Enter 
            Str_Mensagem = Str_Mensagem & "Html --> htm" & Str_Enter 
            Str_Mensagem = Str_Mensagem & "Jpeg --> jpg" & Str_Enter 
            Str_Mensagem = Str_Mensagem & "Mpeg --> mpg." & Str_Enter 

        Case "?6"
            Str_Mensagem = "6 = MUDANÇA DE EXTENÇÃO" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Altera todas as extenções xxx para yyy." & Str_Enter 

        Case "?7"
            Str_Mensagem = "7 = CONTAGEM" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Retorna quantidade de arquivos cujo nome possui palavra desejada." & Str_Enter

        Case "?8"
            Str_Mensagem = "8 = RENUMERAÇÃO DE DIRETÓRIOS" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Renomeia os diretórios de acordo com um número inicial." & Str_Enter

        Case "?9"
            Str_Mensagem = "9 = ESPAÇOS PARA UNDERLINE" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Converte todos os espaços para underline." & Str_Enter

        Case "?10"
            Str_Mensagem = "10 = UNDERLINE PARA ESPAÇOS" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Converte todos os underline para espaços." & Str_Enter

        Case "?11"
            Str_Mensagem = "11 = SUBTRAÇÃO NO NOME DE DIRETÓRIOS" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Retira uma palavra do nome do diretório." & Str_Enter

        Case "?12"
            Str_Mensagem = "12 = SUBSTITUIÇÃO DE SUBSTRING" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "1)Através de uma substring seleciona os arquivos que serão afetados." & Str_Enter 
            Str_Mensagem = Str_Mensagem & "2)Pergunta a substring que será substituída." & Str_Enter 
            Str_Mensagem = Str_Mensagem & "3)Pergunta a nova substring que substuirá a antiga." & Str_Enter 
            
        Case "?13"
            Str_Mensagem = "13 = ALTERAÇÃO DE CAIXA" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Converte caixa (case) para o formato desejado." & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "NOMES EM CAIXA ALTA (A)" & Str_Enter
            Str_Mensagem = Str_Mensagem & "nomes em caixa baixa (B)" & Str_Enter
            Str_Mensagem = Str_Mensagem & "Nomes Com A Primeira Letra Em Caixa Alta (C)" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Obs.: Talvez precise de um refresh (F5) para visualizar resultados." & Str_Enter
        
        Case "?14"
            Str_Mensagem = "14 = Tumblr Clean" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Retira os identificadores do Tumblr" & Str_Enter & Str_Enter
    
        Case "?15"
            Str_Mensagem = "15 = Inclusão de prefixo" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Inclui um prefixo em TODOS os arquivos" & Str_Enter & Str_Enter
    
        Case "?16"
            Str_Mensagem = "16 = Retira acentuação" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Retira todos os acentos do nome dos arquivos" & Str_Enter & str_Enter
        
        Case "?17"
            Str_Mensagem = "17 = Rename Múltiplo Arqs W.Explorer" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Finaliza Rename Múltiplo do Windows Explorer" & Str_Enter & str_Enter
    
        Case "?18"
            Str_Mensagem = "18 = Rename Múltiplo Dirs W.Explorer" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Finaliza Rename Múltiplo de diretórios do Windows Explorer" & Str_Enter & str_Enter
			
        Case "?19"
            Str_Mensagem = "19 = SUBSTITUIÇÃO PARCIAL DIRETÓRIOS" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "1)Através de uma substring seleciona os diretórios que serão afetados." & Str_Enter 
            Str_Mensagem = Str_Mensagem & "2)Pergunta a nova substring que substuirá a nome do diretório." & Str_Enter 
    
        Case "?20"
            Str_Mensagem = "20 = LEFT TRIM" & Str_Enter & Str_Enter
            Str_Mensagem = Str_Mensagem & "Elimina tudo o que estiver a esquerda da palavra fornecida." & Str_Enter 
    
    End Select

    Call msgbox(Str_Mensagem, vbSystemModal + VBInformation, "Descrição da opção " & Replace(Str_opcao,"?", ""))

End Sub

'Monta mensagem de opções do menu
Private Function MensagemMenu

    Str_Mensagem = "ESCOLHA A OPÇÃO DESEJADA:" & Str_Enter 
	Str_Mensagem = Str_Mensagem & " ——————————————— Arquivos ———————————————" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 1 = Zero à esquerda" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 2 = Substituição total" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 3 = Substituição parcial" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 4 = Subtração" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 5 = Simplificação de extenção" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 6 = Mudança de extenção" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 7 = Contagem" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 9 = Espaços para underline" & Str_Enter
    Str_Mensagem = Str_Mensagem & "10 = Underline para espaços" & Str_Enter
    Str_Mensagem = Str_Mensagem & "12 = Substituição de substring" & Str_Enter
    Str_Mensagem = Str_Mensagem & "13 = Alteração de caixa" & Str_Enter
    Str_Mensagem = Str_Mensagem & "14 = Limpeza de Tumblr" & Str_Enter
    Str_Mensagem = Str_Mensagem & "15 = Inclusão de prefixo" & Str_Enter
	Str_Mensagem = Str_Mensagem & "16 = Retira acentuação" & Str_Enter
	Str_Mensagem = Str_Mensagem & "17 = Finaliza Rename Arquivos W.Explorer" & Str_Enter 
	Str_Mensagem = Str_Mensagem & "20 = Left Trim" & Str_Enter 
	Str_Mensagem = Str_Mensagem & Str_Enter & " ——————————————— Diretórios ———————————————" & Str_Enter
    Str_Mensagem = Str_Mensagem & " 8 = Renumeração de diretórios" & Str_Enter
    Str_Mensagem = Str_Mensagem & "11 = Subtração no nome de diretórios" & Str_Enter
	Str_Mensagem = Str_Mensagem & "18 = Finaliza Rename Folders W.Explorer" & Str_Enter
	Str_Mensagem = Str_Mensagem & "19 = Substituição parcial de diretórios" & Str_Enter & Str_Enter
    Str_Mensagem = Str_Mensagem & "ESC = Sair do script" & Str_Enter
    Str_Mensagem = Str_Mensagem & Str_Enter & "? + opção = Ajuda"
    
End Function

'Verifica se arquivo existe
Private Function ArquivoExiste(NomeArquivo)
    If Obj_FileSystem.FileExists(Str_MeuPath & NomeArquivo) Then
        'Somente avisa se quantidade não for grande
        If Int_QtdArquivosFolder < Int_QtdArquivosLimiteAviso Then
            Str_Mensagem = "O arquivo " & NomeArquivo  & " já existe!"
            Call msgbox(Str_Mensagem, VBCritical, "Arquivo existente")
        End If
        ArquivoExiste = True
    Else
        ArquivoExiste = False
    End If
End Function


'Verifica se folder existe
Private Function FolderExiste(NomeFolder)
    If Obj_FileSystem.FolderExists(Str_MeuPath & NomeFolder) Then
		Str_Mensagem = "O arquivo " & NomeFolder  & " já existe!"
		Call msgbox(Str_Mensagem, VBCritical, "Arquivo existente")
        FolderExiste = True
    Else
        FolderExiste = False
    End If
End Function



Private Sub AlertaInfo(pStr_Mensagem)
    Call msgbox(pStr_Mensagem, VBInformation, "Renomeador de arquivos")
End Sub

Private Function CapitalizaNome(pStr_Nome)
    pStr_Nome = LCase(pStr_Nome)
    'Separa as partes do nome (separadas por espaço)
    Arr_PartesNome = Split(pStr_Nome, " ")
    Str_NomeFinal = ""
    dim parteCorrente
    For Int_Indice = 0 to UBound(Arr_PartesNome) 
        parteCorrente = Arr_PartesNome(Int_Indice)
        'adiciona o primeiro caractere da parte em caixa alta
        Str_NomeFinal = Str_NomeFinal & UCase(Left(parteCorrente,1))
        'adiciona o resto da parte em caixa baixa
        Str_NomeFinal = Str_NomeFinal & LCase(Mid(parteCorrente,2))
        'Acrescenta um espaço (separador original)
        Str_NomeFinal = Str_NomeFinal & " "
    Next

    CapitalizaNome = Str_NomeFinal
    
End Function

Private Function RetiraAcento(strTexto)
    strTexto = Replace(strTexto, "á", "a")
    strTexto = Replace(strTexto, "à", "a")
    strTexto = Replace(strTexto, "ã", "a")
    strTexto = Replace(strTexto, "ä", "a")
    strTexto = Replace(strTexto, "â", "a")
    strTexto = Replace(strTexto, "å", "a")
    strTexto = Replace(strTexto, "Á", "A")
    strTexto = Replace(strTexto, "À", "A")
    strTexto = Replace(strTexto, "Ã", "A")
    strTexto = Replace(strTexto, "Ä", "A")
    strTexto = Replace(strTexto, "Â", "A")
    strTexto = Replace(strTexto, "Å", "A")

    strTexto = Replace(strTexto, "è", "e")
    strTexto = Replace(strTexto, "é", "e")
    strTexto = Replace(strTexto, "ê", "e")
    strTexto = Replace(strTexto, "ë", "e")
    strTexto = Replace(strTexto, "È", "E")
    strTexto = Replace(strTexto, "É", "E")
    strTexto = Replace(strTexto, "Ê", "E")
    strTexto = Replace(strTexto, "Ë", "E")
    
    strTexto = Replace(strTexto, "ì", "i")
    strTexto = Replace(strTexto, "í", "i")
    strTexto = Replace(strTexto, "ï", "i")
    strTexto = Replace(strTexto, "î", "i")
    strTexto = Replace(strTexto, "Ì", "I")
    strTexto = Replace(strTexto, "Í", "I")
    strTexto = Replace(strTexto, "Ï", "I")
    strTexto = Replace(strTexto, "Î", "I")
    
    strTexto = Replace(strTexto, "ò", "o")
    strTexto = Replace(strTexto, "ó", "o")
    strTexto = Replace(strTexto, "ô", "o")
    strTexto = Replace(strTexto, "õ", "o")
    strTexto = Replace(strTexto, "ö", "o")
    strTexto = Replace(strTexto, "Ò", "O")
    strTexto = Replace(strTexto, "Ó", "O")
    strTexto = Replace(strTexto, "Ô", "O")
    strTexto = Replace(strTexto, "Õ", "O")
    strTexto = Replace(strTexto, "Ö", "O")

    strTexto = Replace(strTexto, "ù", "u")
    strTexto = Replace(strTexto, "ú", "u")
    strTexto = Replace(strTexto, "û", "u")
    strTexto = Replace(strTexto, "ü", "u")
    strTexto = Replace(strTexto, "Ù", "U")
    strTexto = Replace(strTexto, "Ú", "U")
    strTexto = Replace(strTexto, "Û", "U")
    strTexto = Replace(strTexto, "Ü", "U")
    
    strTexto = Replace(strTexto, "ç", "c")
    strTexto = Replace(strTexto, "Ç", "C")
    
    RetiraAcento = strTexto
End Function

'Monta lista de path Completo dos arquivos 
Private Function MontaListaArquivos()
	Str_ListaNomes = ""
	For Each Obj_Arquivo in Obj_Arquivos
		If Obj_Arquivo.name <> Str_NomeDesteScript Then
			Str_ListaNomes = Str_ListaNomes & Obj_EsteFolder.Path & "\" & Obj_Arquivo.name & "£"		
		End If
	Next
	'Retira o último separador
	Str_ListaNomes = Left(Str_ListaNomes, Len(Str_ListaNomes) - 1)
	Arr_NomesArquivos = Split(Str_ListaNomes, "£")
End Function








