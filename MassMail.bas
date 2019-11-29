Attribute VB_Name = "MassMail"
Option Explicit
Public Sub SendMails()
Dim TheDraft As MailItem
    Set TheDraft = FindDraft("Appel aux cotisations 2019")
    Dim addresses As String, address As Variant
   
    addresses = "catherine.l.aebischer@gmail.com;jessy.ainas@hotmail.fr;danielle-barro@vtxnet.ch;isabatist@gmail.com;baumg@iprolink.ch" _
    + ";florybel@bluewin.ch;kevin.adrien.victor.bigler@gmail.com;lbigwood@vtxnet.ch;jean-marc.binet@sunrise.ch;chrisboreatti@hotmail.com" _
    + ";brugger.manon@bluewin.ch;lise.brulhart@gmail.com;cambitsis.joan@bluewin.ch;cfcarrel@msn.com;dominique.chatelain@bluewin.ch" _
    + ";virginie.chaves-vischer@bluewin.ch;chiuvea@gmail.com;acourvoi@atco.ch;christine.decoulon@deckpoint.ch;jeanpaul.depraz@sunrise.ch" _
    + ";mireille.descombes@gmail.com;eva.dishon@gmail.com;domaretskiy@phystech.edu;christiane.doret@wanadoo.fr;anneduerr@yahoo.com" _
    + ";emmanuelle.cadas@orange.fr;p.feuchere@orange.fr;afoussat@hotmail.fr;ghisleo@yahoo.com;andreag@sfnet.ch" _
    + ";juliette.gilgen@bluewin.ch;bernard.giovannini@unige.ch;paulinegosselin@gmail.com;jmgottraux@bluewin.ch;dominique.haenni@unige.ch" _
    + ";jean-claude.hausmann@unige.ch;collingwood62@me.com;kathrein.humbel@gmail.com;veronika.janjic@bluewin.ch;sandrinejoly@bluewin.ch" _
    + ";h.s.jupp@bluewin.ch;choudans434@gmail.com;milcam@bluewin.ch;ngamauris@bluewin.ch;g.meilhanstalder@vtx.ch" _
    + ";william.meredith@gmail.com;miyaitakuma@live.jp;matthis.pasche@gmail.com;vincent.pilloux@orange.fr;anne.piotet@gmail.com" _
    + ";philippe_raemy@swissonline.ch;heloiserae@gmail.com;severine.laliveraemy@swissonline.ch;arianamorris@hotmail.com;c.ronzaud@yahoo.com" _
    + ";jacques.rougemont@unige.ch;francois.rudhard@wanadoo.fr;melanie.saul@etu.unige.ch;osorg@bluewin.ch;annedospertiniaechallens@gmail.com" _
    + ";gabriela.tappolet@edu.ge.ch;danielle.thien@gmail.com;jvdm@vdmsolutions.com;dirk.vandermarel@unige.ch;michele.vanhersel@orange.fr" _
    + ";julie.veron.749@gmail.com;anne-villeneuve@hotmail.fr;sjvilleneuve@bluewin.ch;fabiennevuilleumier@gmail.com;tomo_cw@yahoo.co.jp" _
    + ";janet.weiler@bluewin.ch;physiowismer@gmx.ch;heinz.zimmermann@sunrise.ch"


    For Each address In Split(addresses, ";")
        If Not address = "" Then
            SendMail TheDraft, CStr(address)
        End If
    Next address
    
End Sub

Public Sub SendMail(ByVal TheDraft As MailItem, address As String)
Dim theMail As MailItem
    
    Set theMail = TheDraft.Copy
    theMail.Recipients.add address
    ThisOutlookSession.disableAutoClassify = True
    theMail.Send
    ThisOutlookSession.disableAutoClassify = False
End Sub



Public Function FindDraft(subject As String) As MailItem
Dim drafts As Folder
Dim item
Dim mi As MailItem

    Set drafts = Application.Session.GetDefaultFolder(olFolderDrafts)
    For Each item In drafts.Items
        If TypeName(item) = "MailItem" Then
            Set mi = item
            If mi.subject = subject Then
                Set FindDraft = mi
                Exit Function
            End If
        End If
    Next item
        
End Function


