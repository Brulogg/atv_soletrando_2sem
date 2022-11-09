dim palavra(20),acerto,nivel,resposta,np,ouvir
dim resp,audio,nome,pular
call carregar_audio

sub carregar_audio()
set audio=createobject("SAPI.SPVOICE")
audio.volume = 100
audio.rate = 2 'Velocidade da fala
call carregar_nome
end sub

sub carregar_nome()
nome=inputbox("Insira seu Nickname: ")
call carregar_pre
end sub

sub carregar_pre()
acerto=0
nivel=1
pular=0
call carregar_palavra
end sub

sub carregar_palavra()
palavra(0)="peixe"
palavra(1)="bola"
palavra(2)="gato"
palavra(3)="verde"
palavra(4)="rosa"
palavra(5)="dezena"
palavra(6)="vermelho"
palavra(7)="mamão"
palavra(8)="sócio"
palavra(9)="cérebro"
palavra(10)="piscina"
palavra(11)="incompatibilidade"
palavra(12)="isenção"
palavra(13)="repreensão"
palavra(14)="exceção"
palavra(15)="réveillon"
palavra(16)="zooplâncton"
palavra(17)="otorrinolaringologista"
palavra(18)="inconstitucionalissimamente"
palavra(19)="impeachment"
call carregar_nivel1
end sub

sub carregar_nivel1()
do while nivel=1
    ouvir=0
    randomize(second(time))
    np=int(rnd * 5)
    if palavra(np)="false" then
        np=int(rnd * 5)
    else 
        audio.speak ("A palavra é: "&palavra(np)&"")
        resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                        "[O]uvir Novamente"+vbnewline &_
                        "[P]ular uma única vez"+vbnewline &_
                        "Nível: "& nivel &""))
        if resposta="o" then
            if ouvir=0 then
                audio.speak ("A palavra é: "&palavra(np)&"")
                do while resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                    "[O]uvir Novamente"+vbnewline &_
                                    "[P]ular uma única vez"+vbnewline &_
                                 "Nível: "& nivel &""))
                    ouvir=1
                    if resposta="o" then
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            else
                msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
            end if
        end if
        if resposta="p" then
            if pular=0 then
                palavra(np)="false"
                pular=1
            else
                msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                do while resposta="p" or resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                        "[O]uvir Novamente"+vbnewline &_
                                        "[P]ular uma única vez"+vbnewline &_
                                        "Nível: "& nivel &""))
                    if resposta="o" then
                        if ouvir=0 then
                            audio.speak ("A palavra é: "&palavra(np)&"")
                            ouvir=1
                        else
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                        end if
                    end if
                    if resposta="p" then
                        msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            end if
        end if
        if resposta<>palavra(np) and resposta<>"p" and resposta<>"o" then
            resp=msgbox("Você Errou!"+vbnewline &_
                    "Acertos: "&acerto&""+vbnewline &_
                    "Nível: "&nivel&""+vbnewline &_
                    "Palavra: "&palavra(np)&""+vbnewline &_
                    "Deseja jogar novamente?",vbquestion+vbyesno,"AVISO")
            if resp=vbyes then
                call carregar_pre
            else
                wscript.quit
            end if
        end if
        if resposta=palavra(np) then
            msgbox("Você Acertou!"),vbinformation+vbokonly,"AVISO"
            acerto=acerto+1
            palavra(np)="false"
        end if
        if acerto=3 then
            nivel=2
            call carregar_nivel2
        end if
    end if
loop
end sub

sub carregar_nivel2()
msgbox("Você passou para o nivel 2"),vbinformation+vbokonly,"AVISO"
do while nivel=2
    ouvir=0
    randomize(second(time))
    np=int(rnd * 5)+5
    if palavra(np)="false" then
        np=int(rnd * 5)+5
    else 
        audio.speak ("A palavra é: "&palavra(np)&"")
        resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                        "[O]uvir Novamente"+vbnewline &_
                        "[P]ular uma única vez"+vbnewline &_
                        "Nível: "& nivel &""))
        if resposta="o" then
            if ouvir=0 then
                audio.speak ("A palavra é: "&palavra(np)&"")
                do while resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                    "[O]uvir Novamente"+vbnewline &_
                                    "[P]ular uma única vez"+vbnewline &_
                                 "Nível: "& nivel &""))
                    ouvir=1
                    if resposta="o" then
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            else
                msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
            end if
        end if
        if resposta="p" then
            if pular=0 then
                palavra(np)="false"
                pular=1
            else
                msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                do while resposta="p" or resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                        "[O]uvir Novamente"+vbnewline &_
                                        "[P]ular uma única vez"+vbnewline &_
                                        "Nível: "& nivel &""))
                    if resposta="o" then
                        if ouvir=0 then
                            audio.speak ("A palavra é: "&palavra(np)&"")
                            ouvir=1
                        else
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                        end if
                    end if
                    if resposta="p" then
                        msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            end if
        end if
        if resposta<>palavra(np) and resposta<>"p" and resposta<>"o" then
            resp=msgbox("Você Errou!"+vbnewline &_
                    "Acertos: "&acerto&""+vbnewline &_
                    "Nível: "&nivel&""+vbnewline &_
                    "Palavra: "&palavra(np)&""+vbnewline &_
                    "Deseja jogar novamente?",vbquestion+vbyesno,"AVISO")
            if resp=vbyes then
                call carregar_pre
            else
                wscript.quit
            end if
        end if
        if resposta=palavra(np) then
            msgbox("Você Acertou!"),vbinformation+vbokonly,"AVISO"
            acerto=acerto+1
            palavra(np)="false"
        end if
        if acerto=6 then
            nivel=3
            call carregar_nivel3
        end if
    end if
loop
end sub

sub carregar_nivel3()
msgbox("Você passou para o nivel 3"),vbinformation+vbokonly,"AVISO"
do while nivel=3
    ouvir=0
    randomize(second(time))
    np=int(rnd * 5)+10
    if palavra(np)="false" then
        np=int(rnd * 5)+10
    else 
        audio.speak ("A palavra é: "&palavra(np)&"")
        resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                        "[O]uvir Novamente"+vbnewline &_
                        "[P]ular uma única vez"+vbnewline &_
                        "Nível: "& nivel &""))
        if resposta="o" then
            if ouvir=0 then
                audio.speak ("A palavra é: "&palavra(np)&"")
                do while resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                    "[O]uvir Novamente"+vbnewline &_
                                    "[P]ular uma única vez"+vbnewline &_
                                 "Nível: "& nivel &""))
                    ouvir=1
                    if resposta="o" then
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            else
                msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
            end if
        end if
        if resposta="p" then
            if pular=0 then
                palavra(np)="false"
                pular=1
            else
                msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                do while resposta="p" or resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                        "[O]uvir Novamente"+vbnewline &_
                                        "[P]ular uma única vez"+vbnewline &_
                                        "Nível: "& nivel &""))
                    if resposta="o" then
                        if ouvir=0 then
                            audio.speak ("A palavra é: "&palavra(np)&"")
                            ouvir=1
                        else
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                        end if
                    end if
                    if resposta="p" then
                        msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            end if
        end if
        if resposta<>palavra(np) and resposta<>"p" and resposta<>"o" then
            resp=msgbox("Você Errou!"+vbnewline &_
                    "Acertos: "&acerto&""+vbnewline &_
                    "Nível: "&nivel&""+vbnewline &_
                    "Palavra: "&palavra(np)&""+vbnewline &_
                    "Deseja jogar novamente?",vbquestion+vbyesno,"AVISO")
            if resp=vbyes then
                call carregar_pre
            else
                wscript.quit
            end if
        end if
        if resposta=palavra(np) then
            msgbox("Você Acertou!"),vbinformation+vbokonly,"AVISO"
            acerto=acerto+1
            palavra(np)="false"
        end if
        if acerto=9 then
            nivel=4
            call carregar_nivel4
        end if
    end if
loop
end sub

sub carregar_nivel4()
msgbox("Você passou para o nivel 4"),vbinformation+vbokonly,"AVISO"
do while nivel=4
    ouvir=0
    randomize(second(time))
    np=int(rnd * 5)+15
    if palavra(np)="false" then
        np=int(rnd * 5)+15
    else 
        audio.speak ("A palavra é: "&palavra(np)&"")
        resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                        "[O]uvir Novamente"+vbnewline &_
                        "[P]ular uma única vez"+vbnewline &_
                        "Nível: "& nivel &""))
        if resposta="o" then
            if ouvir=0 then
                audio.speak ("A palavra é: "&palavra(np)&"")
                do while resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                    "[O]uvir Novamente"+vbnewline &_
                                    "[P]ular uma única vez"+vbnewline &_
                                 "Nível: "& nivel &""))
                    ouvir=1
                    if resposta="o" then
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            else
                msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
            end if
        end if
        if resposta="p" then
            if pular=0 then
                palavra(np)="false"
                pular=1
            else
                msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                do while resposta="p" or resposta="o"
                    resposta=LCase(inputbox("Nick: "& nome &""+vbnewline &_
                                        "[O]uvir Novamente"+vbnewline &_
                                        "[P]ular uma única vez"+vbnewline &_
                                        "Nível: "& nivel &""))
                    if resposta="o" then
                        if ouvir=0 then
                            audio.speak ("A palavra é: "&palavra(np)&"")
                            ouvir=1
                        else
                        msgbox("Você não pode ouvir novamente"),vbinformation+vbokonly,"AVISO"
                        end if
                    end if
                    if resposta="p" then
                        msgbox("Você não pode pular"),vbinformation+vbokonly,"AVISO"
                    end if
                loop
            end if
        end if
        if resposta<>palavra(np) and resposta<>"p" and resposta<>"o" then
            resp=msgbox("Você Errou!"+vbnewline &_
                    "Acertos: "&acerto&""+vbnewline &_
                    "Nível: "&nivel&""+vbnewline &_
                    "Palavra: "&palavra(np)&""+vbnewline &_
                    "Deseja jogar novamente?",vbquestion+vbyesno,"AVISO")
            if resp=vbyes then
                call carregar_pre
            else
                wscript.quit
            end if
        end if
        if resposta=palavra(np) then
            msgbox("Você Acertou!"),vbinformation+vbokonly,"AVISO"
            acerto=acerto+1
            palavra(np)="false"
        end if
        if acerto=12 then
            nivel=5
            resp=msgbox("Parabéns! Você Ganhou!!"+vbnewline &_
                        "Acertos: "&acerto&""+vbnewline &_
                        "Nível: Final"+vbnewline &_
                        "Deseja jogar novamente?",vbquestion+vbyesno,"PARABENS")
            if resp=vbyes then
                call carregar_pre
            else
                wscript.quit
            end if
        end if
    end if
loop
end sub