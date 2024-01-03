Attribute VB_Name = "Client_UpdateWindow"
Public Sub UpdateWindowOffer(ByVal Index_Offer As Long)
    Dim i As Long
    ' gui stuff
    With Windows(GetWindowIndex("winOffer"))
        ' set main text
        If Index_Offer <> 0 Then
            .Controls(GetControlIndex("winOffer", "picBGOffer" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "picOfferBG" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "btnAccept" & Index_Offer)).visible = True
            .Controls(GetControlIndex("winOffer", "btnRecuse" & Index_Offer)).visible = True
            Select Case inOfferType(Index_Offer)
                Case Offers.Offer_Type_Party
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = inOfferInvite(Index_Offer) & " has invited you to a party."
                Case Offers.Offer_Type_Trade
                    .Controls(GetControlIndex("winOffer", "lblTitleOffer" & Index_Offer)).text = inOfferInvite(Index_Offer) & "  has invited you to trade."
            End Select
            ShowWindow GetWindowIndex("winOffer")
        Else
            For i = 1 To MAX_OFFER
                .Controls(GetControlIndex("winOffer", "picBGOffer" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "picOfferBG" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "lblTitleOffer" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "btnAccept" & i)).visible = False
                .Controls(GetControlIndex("winOffer", "btnRecuse" & i)).visible = False
            Next
            HideWindow GetWindowIndex("winOffer")
        End If
    End With
End Sub
