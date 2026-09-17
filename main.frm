VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm main 
   BackColor       =   &H008F4100&
   Caption         =   "data_pro"
   ClientHeight    =   9930
   ClientLeft      =   165
   ClientTop       =   -645
   ClientWidth     =   13260
   LinkTopic       =   "MDIForm1"
   Picture         =   "main.frx":0000
   RightToLeft     =   -1  'True
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin Crystal.CrystalReport REPORT1 
      Left            =   2700
      Top             =   180
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Menu mnitems 
      Caption         =   "√’‰«›"
      NegotiatePosition=   1  'Left
      Begin VB.Menu tmmodels 
         Caption         =   "»Ì«‰«  «·„ÊœÌ·« "
      End
      Begin VB.Menu tmfixcostsales 
         Caption         =   "Ÿ»ÿ  ﬂ·›… «·„»Ì⁄« "
      End
      Begin VB.Menu tmdataitem 
         Caption         =   "»Ì«‰«   ›’Ì·Ï ··„ÊœÌ·« "
      End
      Begin VB.Menu tmmodelokazgr 
         Caption         =   "„Ã„Ê⁄«  ⁄—Ê÷ Œ’„ «·„ÊœÌ·« "
      End
      Begin VB.Menu tmitemgroup 
         Caption         =   "„Ã„Ê⁄«  „ÊœÌ·« "
      End
      Begin VB.Menu tmgrmain 
         Caption         =   "„Ã„Ê⁄«  —∆Ì”Ì…"
      End
      Begin VB.Menu tmstores 
         Caption         =   "»Ì«‰«  «·„Œ«“‰"
      End
      Begin VB.Menu tmfact 
         Caption         =   "»Ì«‰«  „’«‰⁄"
      End
      Begin VB.Menu tmsec 
         Caption         =   "«·√ﬁ”«„"
      End
      Begin VB.Menu tmsexcode 
         Caption         =   "«·‰Ê⁄"
      End
      Begin VB.Menu tmagecode 
         Caption         =   "«·›∆… «·⁄„—Ì…"
      End
      Begin VB.Menu tmtrans 
         Caption         =   " ÕÊÌ·«  «’‰«›"
      End
      Begin VB.Menu trmtrans2 
         Caption         =   " ÕÊÌ· ’«œ— „‰ «·›—⁄"
      End
      Begin VB.Menu tmtranscost 
         Caption         =   "≈Ã„«·Ï ﬁÌ„…  ÕÊÌ·«  »Ì‰ «·›—Ê⁄"
      End
      Begin VB.Menu rmdocstock 
         Caption         =   "≈Ã„«·Ï Ã—œ „Œ«“‰"
      End
      Begin VB.Menu tmstock 
         Caption         =   "Ã—œ „Œ«“‰"
      End
      Begin VB.Menu tmstockitemdoc 
         Caption         =   " ›’Ì·Ï „” ‰œ«  Ã—œ «’‰«› "
      End
      Begin VB.Menu xItemMove 
         Caption         =   "Õ—ﬂ… „ÊœÌ·"
      End
      Begin VB.Menu tmtransbarcode 
         Caption         =   " ⁄œÌ· »«—ﬂÊœ ··›—Ê⁄"
      End
   End
   Begin VB.Menu mnclients 
      Caption         =   "⁄„·«¡"
      Begin VB.Menu tmcust 
         Caption         =   "»Ì«‰«  ⁄„·«¡ √Ã·"
      End
      Begin VB.Menu tmmovecust 
         Caption         =   "Õ—ﬂ… ⁄„·«¡ √Ã·"
      End
      Begin VB.Menu tm_grid_inv_balance 
         Caption         =   " ﬁ—Ì— «—’œ… «·›Ê« Ì—"
      End
      Begin VB.Menu tm_list_price 
         Caption         =   "ﬁ«∆„… «”⁄«— «·⁄„·«¡"
      End
   End
   Begin VB.Menu mnVendorsx 
      Caption         =   "„Ê—œÌ‰"
      Begin VB.Menu tmsuppl 
         Caption         =   "»Ì«‰«  „Ê—œÌ‰"
      End
      Begin VB.Menu tmsuppgr 
         Caption         =   "„Ã„Ê⁄«  „Ê—œÌ‰"
      End
      Begin VB.Menu tmtransmosm 
         Caption         =   " —ÕÌ· Õ”«»«  „Ê«”„"
      End
      Begin VB.Menu tmsuppmove 
         Caption         =   "Õ—ﬂ… „Ê—œÌ‰ - „Ê«”„"
      End
      Begin VB.Menu tmmovesupp 
         Caption         =   "Õ—ﬂ… «·„Ê—œÌ‰"
      End
   End
   Begin VB.Menu mnInvoice 
      Caption         =   "›Ê« Ì—"
      Begin VB.Menu tmsales 
         Caption         =   "„»Ì⁄« "
      End
      Begin VB.Menu tm_sales_orders 
         Caption         =   "⁄—Ê÷ «”⁄«—"
      End
      Begin VB.Menu tmDashBoard 
         Caption         =   "DashBoard"
      End
      Begin VB.Menu tmTargetBr 
         Caption         =   "„” Âœ› «·„»Ì⁄«  «·‘Â—Ï ··›—Ê⁄"
      End
      Begin VB.Menu tmreptarget 
         Caption         =   " ﬁ—Ì— «·„” Âœ› ·›—⁄ ·› —…"
      End
      Begin VB.Menu tmRepMonthTot 
         Caption         =   "≈Ã„«·Ï „»Ì⁄«  & „” Âœ› ··›—Ê⁄"
      End
      Begin VB.Menu tmSalesBranch 
         Caption         =   "„»Ì⁄«  «·›—Ê⁄"
      End
      Begin VB.Menu tmopensal 
         Caption         =   "»Ê‰«   ”·Ì„ «Ê‰ ·«Ì‰"
      End
      Begin VB.Menu tmsalestax 
         Caption         =   "„»Ì⁄«  ›Ê« Ì— ÷—Ì»Ì»…"
      End
      Begin VB.Menu tmsalesinvret 
         Caption         =   "„—œÊœ „»Ì⁄«  ÷—Ì»Ì»…"
      End
      Begin VB.Menu tmrepinvoice 
         Caption         =   " ﬁ—Ì— ≈Ã„«·Ï ›Ê« Ì— ÷—Ì»Ì…"
      End
      Begin VB.Menu tmorderamazon 
         Caption         =   "ÿ·»Ì«  ⁄„·«¡ Ã„·…"
      End
      Begin VB.Menu tmRecevdOrder 
         Caption         =   "≈” ·«„ „ÊœÌ·«  ·ÿ·»Ì… ⁄„·«¡"
      End
      Begin VB.Menu tmorder_rec_amazon 
         Caption         =   " ﬁ—Ì— „ﬁ«—‰… ÿ·»Ì… - ≈” ·«„ - „»Ì⁄«  - ÿ·»Ì«  ⁄„·«¡"
      End
      Begin VB.Menu tmrecordercust 
         Caption         =   "≈Ã„«·Ï ÿ·»Ì«  -  ”·Ì„ - „»Ì⁄« "
      End
      Begin VB.Menu tmonline_order 
         Caption         =   "ÿ·»Ì«  «Ê‰ ·«Ì‰"
      End
      Begin VB.Menu tmpayshiponline 
         Caption         =   "”œ«œ »Ê«·’ ‘Õ‰ «Ê‰·«Ì‰"
      End
      Begin VB.Menu tmdashonline 
         Caption         =   "dashboard ONLINE"
      End
      Begin VB.Menu tmorderreponline 
         Caption         =   " ﬁ—Ì— „Ã„⁄ ÿ·»Ì«  «Ê‰·«Ì‰"
      End
      Begin VB.Menu tmpurchases 
         Caption         =   "„‘ —Ì« "
      End
      Begin VB.Menu tmRETPURCH 
         Caption         =   "„—œÊœ „‘ —Ì« "
      End
      Begin VB.Menu tmdocsend 
         Caption         =   "„—«Ã⁄… «” ·«„ «·„” ‰œ«  ··›—Ê⁄"
      End
      Begin VB.Menu tmNewDocBranch 
         Caption         =   "„—«Ã⁄… «” ·«„ „” ‰œ«  ›—⁄ «·”Ì—›—"
      End
      Begin VB.Menu tmrepdemorder 
         Caption         =   "ÿ·»Ì«  «·›—Ê⁄ „Ê“⁄… „Ê—œÌ‰ - „Ê«”„ - «ﬁ”«„"
      End
      Begin VB.Menu tmorderpurch 
         Caption         =   "ÿ·»«  „‘ —Ì«  "
      End
      Begin VB.Menu tmrecorder_supler 
         Caption         =   "≈” ·«„ ÿ·»Ì«  „‰ «·„Ê—œÌ‰"
      End
      Begin VB.Menu tmpurchordertot 
         Caption         =   "„ «»⁄… ≈Ã„«·Ï ÿ·»«  „‘ —Ì«  „‰ «·„Ê—œÌ‰"
      End
      Begin VB.Menu tmOrderDem 
         Caption         =   "„ «»⁄…  ›’Ì·Ï ÿ·»Ì«  „Ê—œÌ‰"
      End
      Begin VB.Menu tmOrderDem_ALL 
         Caption         =   "„ﬁ«—‰…  ›’Ì·Ì… ÿ·»Ì«  „ÊœÌ·«  Ê „‘ —Ì«  „ÊœÌ·« "
      End
      Begin VB.Menu tmOrderDOC_ALL 
         Caption         =   "„ﬁ«—‰… ≈Ã„«·Ï  ÿ·»Ì«  Ê „‘ —Ì«  "
      End
      Begin VB.Menu line51 
         Caption         =   "-"
      End
      Begin VB.Menu tmbarcode 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
      End
      Begin VB.Menu tmrepvisa 
         Caption         =   "„—«Ã⁄… „»Ì⁄«   «·›Ì“«"
      End
      Begin VB.Menu tmsalday 
         Caption         =   "„—«Ã⁄… «·„»Ì⁄«  «·›—Ê⁄"
      End
      Begin VB.Menu tmret_sales 
         Caption         =   " ﬁ—Ì— „— Ã⁄«  «·»Ê‰« "
      End
      Begin VB.Menu tmmansales2 
         Caption         =   "„—«Ã⁄… „»Ì⁄«  «·»«∆⁄Ì‰"
      End
   End
   Begin VB.Menu mn_online 
      Caption         =   "«Ê‰·«Ì‰"
      Begin VB.Menu tm_online_orders 
         Caption         =   "ÿ·»Ì«  «Ê‰ ·«Ì‰"
      End
   End
   Begin VB.Menu mn_MAINCash 
      Caption         =   "‰ﬁœÌ…"
      Begin VB.Menu xCASH2 
         Caption         =   "„œ›Ê⁄«  «·Ì «·„Ê—œÌ‰"
      End
      Begin VB.Menu tmcash0 
         Caption         =   " ”ÊÌ«  „Ê—œÌ‰"
      End
      Begin VB.Menu xmcash3 
         Caption         =   "„ﬁ»Ê÷«  „‰ «·„Ê—œÌ‰"
      End
      Begin VB.Menu xCashed 
         Caption         =   "„ﬁ»Ê÷«  „‰ ⁄„·«¡"
      End
      Begin VB.Menu xmcash4 
         Caption         =   "„œ›Ê⁄«  «·Ì «·⁄„·«¡"
      End
      Begin VB.Menu tmcash11 
         Caption         =   " ”ÊÌ«  ⁄„·«¡"
      End
      Begin VB.Menu tmcharge 
         Caption         =   "„’«—Ì›"
      End
      Begin VB.Menu tmcharge2 
         Caption         =   "„’«—Ì›  ÊÌﬂ·« "
      End
      Begin VB.Menu tmchargecode 
         Caption         =   "«ﬂÊ«œ „’«—Ì›"
      End
      Begin VB.Menu tmchargemaincode 
         Caption         =   "«ﬂÊ«œ „’«—Ì› —∆Ì”Ì…"
      End
      Begin VB.Menu tmcashbranch 
         Caption         =   " ”ÃÌ· ‰ﬁœÌ… «·›—Ê⁄"
      End
      Begin VB.Menu tmincome 
         Caption         =   " ”ÃÌ· «Ì—«œ« "
      End
      Begin VB.Menu tmincomecode 
         Caption         =   "«ﬂÊ«œ «Ì—«œ« "
      End
      Begin VB.Menu tmincomemaincode 
         Caption         =   "«ﬂÊ«œ «Ì—«œ«  —∆Ì”Ì…"
      End
      Begin VB.Menu tmpart 
         Caption         =   "Ã«—Ì «·‘—ﬂ«¡"
      End
      Begin VB.Menu tmpart_code 
         Caption         =   "«ﬂÊ«œ «·‘—ﬂ«¡"
      End
      Begin VB.Menu tmbox 
         Caption         =   "«ﬂÊ«œ Œ“‰"
      End
      Begin VB.Menu tm_box_day 
         Caption         =   "ÌÊ„Ì… Œ“‰"
      End
      Begin VB.Menu tmboxtrans 
         Caption         =   " ÕÊÌ·«  Œ“‰"
      End
      Begin VB.Menu tmtranscode2 
         Caption         =   " ÕÊÌ·«  Œ“‰ «·›—Ê⁄"
      End
      Begin VB.Menu tmboxbal 
         Caption         =   "—’Ìœ «·Œ“‰…"
      End
      Begin VB.Menu tmbalbox 
         Caption         =   "≈Ã„«·Ï —’Ìœ «·Œ“‰"
      End
   End
   Begin VB.Menu mn_Bank 
      Caption         =   "»‰Êﬂ"
      Begin VB.Menu tmBankData 
         Caption         =   "»Ì«‰«  «·»‰Êﬂ"
      End
      Begin VB.Menu tmBankGrroup 
         Caption         =   "„Ã„Ê⁄«  «·»‰Êﬂ"
      End
      Begin VB.Menu tmBankMove 
         Caption         =   "Õ—ﬂ… «·»‰Êﬂ"
      End
      Begin VB.Menu tmBankItems 
         Caption         =   "«ﬂÊ«œ Õ—ﬂ… «·»‰ﬂ"
      End
      Begin VB.Menu tmBankInout 
         Caption         =   " ”ÃÌ· Õ—ﬂ… «·»‰Êﬂ"
      End
      Begin VB.Menu tmacccode 
         Caption         =   "«ﬂÊœ Õ”«»«  «Œ—Ï"
      End
      Begin VB.Menu tmtransacc 
         Caption         =   " ”ÊÌ… Õ”«»« "
      End
      Begin VB.Menu tmmoveacc 
         Caption         =   "ﬂ‘› Õ”«» Õ”«»«  «Œ—Ï"
      End
      Begin VB.Menu tmbalacc 
         Caption         =   "«—’œ… «·Õ”«»«  «·«Œ—Ï"
      End
      Begin VB.Menu tmChqIn 
         Caption         =   "√Ê—«ﬁ ﬁ»÷"
      End
      Begin VB.Menu tmChqOut 
         Caption         =   "√Ê—«ﬁ œ›⁄"
      End
      Begin VB.Menu tmBankState 
         Caption         =   "ﬂ‘› Õ”«» »‰ﬂ"
      End
      Begin VB.Menu tmBankMoveTotal 
         Caption         =   "«Ã„«·Ì Õ—ﬂ… «·»‰ﬂ"
      End
      Begin VB.Menu tmbankDtl1 
         Caption         =   " ﬁ—Ì—  ›’Ì·Ì Õ—ﬂ… «·»‰ﬂ"
      End
   End
   Begin VB.Menu mnCASH 
      Caption         =   " ﬁ«—Ì—"
      Begin VB.Menu tm_vsitem 
         Caption         =   "„ «»⁄… «·„ÊœÌ·«  ··„Õ·« "
      End
      Begin VB.Menu tmbalbrach 
         Caption         =   "„—«Ã⁄… «·«—’œ… ··›—Ê⁄"
      End
      Begin VB.Menu tmbalbrachFR 
         Caption         =   "„—«Ã⁄… «·«—’œ… ·· ÊﬂÌ·« "
      End
      Begin VB.Menu tmitemshop 
         Caption         =   "„ «»⁄… «·„ÊœÌ·«   ·›—⁄"
      End
      Begin VB.Menu tmitemshopITEM 
         Caption         =   "„ «»⁄… «·„ÊœÌ·«    ›’Ì·Ï „ﬁ«” - ·Ê‰ ·›—⁄"
      End
      Begin VB.Menu tmallbranchsals 
         Caption         =   "„ «»⁄… «Ã„«·Ì… ··›—Ê⁄ Ê «· ÊﬂÌ·«  - „‘ —Ì«  - „»Ì⁄«  - —’Ìœ -‰”»… »Ì⁄"
      End
      Begin VB.Menu tmbalstore 
         Caption         =   "—’Ìœ «·„ÊœÌ·«  „Ê“⁄ ⁄·Ï «·›—Ê⁄ Ê  «· ÊﬂÌ·« "
      End
      Begin VB.Menu tmlookitem 
         Caption         =   "≈” ⁄·«„ «—’œ… „ÊœÌ·«  -  ›’Ì·Ï „ﬁ«”« "
      End
      Begin VB.Menu tmbalstore_scal 
         Caption         =   " ›’Ì·Ï —’Ìœ „ﬁ«”«  „ÊœÌ·«  ·›—⁄ «Ê ›—Ê⁄"
      End
      Begin VB.Menu tmnewprice 
         Caption         =   "«”⁄«—  „  ⁄œÌ·Â«"
      End
      Begin VB.Menu tmrepeicefact 
         Caption         =   " ÕœÌÀ «·«”⁄«— „‰ «·„’‰⁄"
      End
      Begin VB.Menu tmmodelfact 
         Caption         =   "»Ì«‰«  «·„ÊœÌ·«  „‰ «·„’‰⁄"
      End
      Begin VB.Menu tmNewDoc 
         Caption         =   "„” ‰œ«  ÃœÌœ… ··›—⁄"
      End
      Begin VB.Menu tmmodelshow 
         Caption         =   "⁄—÷ «·»Ì«‰« "
      End
      Begin VB.Menu tmreorder 
         Caption         =   "⁄—÷ ÿ·»Ì«   ›—Ê⁄ «·Êﬂ·«¡"
      End
      Begin VB.Menu tmInput_ALL_Tab 
         Caption         =   " ﬁ«—Ì— «·Ê«—œ „Ã„⁄…"
      End
      Begin VB.Menu tm_grd_input 
         Caption         =   " ﬁ—Ì— «·Ê«—œ ⁄·Ì „” ÊÌ «·‘ÂÊ—"
      End
      Begin VB.Menu tmrepbal 
         Caption         =   " ﬁ«—Ì— «·—’Ìœ  «·›—Ê⁄ Ê «· ÊﬂÌ·« --"
         Begin VB.Menu tmmosmsecbal 
            Caption         =   "ﬁÌ„… «·—’Ìœ „Ê“⁄ ⁄·Ï «·«ﬁ”«„ Ê «·›—Ê⁄"
         End
         Begin VB.Menu tmmosmgrbal 
            Caption         =   "ﬁÌ„… «·—’Ìœ „Ê“⁄ ⁄·Ï «·„Ê«”„ Ê «·„Ã„Ê⁄« "
         End
         Begin VB.Menu tmgroupsection 
            Caption         =   "≈Ã„«·Ï „Êﬁ› «·„Ã„Ê⁄«  "
         End
         Begin VB.Menu tmcostbalbransh 
            Caption         =   "≈Ã„«·Ï „Êﬁ› «·„’«‰⁄"
         End
         Begin VB.Menu tmallsupp 
            Caption         =   "≈Ã„«·Ï „Êﬁ› «·„Ê—œÌ‰"
         End
         Begin VB.Menu tmallSEC 
            Caption         =   "≈Ã„«·Ï „Êﬁ› «·«ﬁ”«„"
         End
         Begin VB.Menu tmbalage 
            Caption         =   "≈Ã„«·Ï „Êﬁ› «·›∆«  «·⁄„—Ì…"
         End
         Begin VB.Menu tmbalmodelscal 
            Caption         =   " ﬁ—Ì— —’Ìœ „ÊœÌ·«  „Ê“⁄ „ﬁ«”« "
         End
      End
      Begin VB.Menu tm_grdBalancePrice 
         Caption         =   " ﬁÌÌ„ «·—’Ìœ »«·«”⁄«— „Ã„⁄"
      End
      Begin VB.Menu tmbalallcost 
         Caption         =   " ﬁÌ„ «·—’Ìœ »«·«”⁄«— - ›—Ê⁄ Ê  ÊﬂÌ·« --"
         Begin VB.Menu tmbalmodelsec 
            Caption         =   " ﬁÌ„ —’Ìœ «·√ﬁ”«„ Ì«”⁄«— „’‰⁄ -  ﬂ·›… -  «·„” Â·ﬂ"
         End
         Begin VB.Menu tmbalmodelgr 
            Caption         =   " ﬁÌ„ —’Ìœ «·„Ã„Ê⁄«  Ì«”⁄«— „’‰⁄ -  ﬂ·›… -  «·„” Â·ﬂ"
         End
         Begin VB.Menu tmbalmodecode 
            Caption         =   " ﬁÌ„ —’Ìœ «·„Ê—œÌ‰  Ì«”⁄«— „’‰⁄ -  ﬂ·›… -  «·„” Â·ﬂ"
         End
         Begin VB.Menu tmbalmodemodel 
            Caption         =   " ﬁÌ„ —’Ìœ «·„ÊœÌ·«  Ì«”⁄«— „’‰⁄ -  ﬂ·›… - «·„” Â·ﬂ"
         End
         Begin VB.Menu tmbaLSTORE_PRICE 
            Caption         =   " ﬁÌ„ —’Ìœ «·›—Ê⁄ Ì«”⁄«— „’‰⁄ -  ﬂ·›… -  «·„” Â·ﬂ"
         End
      End
      Begin VB.Menu tmsalesyear 
         Caption         =   "„ﬁ«—‰… „»Ì⁄«  ”‰Ê« --"
         Begin VB.Menu tmsalessSUPear 
            Caption         =   "„ﬁ«—‰…  „»Ì⁄«   „Ê—œÌ‰ ··”‰Ê« "
         End
         Begin VB.Menu tmsalessecyear 
            Caption         =   "„ﬁ«—‰…  „»Ì⁄«   «ﬁ”«„ ··”‰Ê« "
         End
         Begin VB.Menu tmsalesstoreyear 
            Caption         =   "„ﬁ«—‰…  „»Ì⁄«  «·›—Ê⁄ ·”‰Ê« "
         End
         Begin VB.Menu tmsalesmonth_year 
            Caption         =   "„ﬁ«—‰… „»Ì⁄«  «·‘ÂÊ— ··”‰Ê« "
         End
         Begin VB.Menu tmsalesmonth_MODEL 
            Caption         =   "„ﬁ«—‰… „»Ì⁄«  «·„ÊœÌ·«  ··”‰Ê« "
         End
      End
      Begin VB.Menu tmpURCHyEAR 
         Caption         =   "„ﬁ«—‰… „‘ —Ì«  ”‰Ê« --"
         Begin VB.Menu tmPuechSUPear 
            Caption         =   "„ﬁ«—‰…  „‘ —Ì«   „Ê—œÌ‰ ··”‰Ê« "
         End
         Begin VB.Menu tmPurchSecYear 
            Caption         =   "„ﬁ«—‰…  „‘ —Ì«  «ﬁ”«„ ··”‰Ê« "
         End
         Begin VB.Menu tmPURCHstoreyear 
            Caption         =   "„ﬁ«—‰…  „‘ —Ì«  «·›—Ê⁄ ·”‰Ê« "
         End
         Begin VB.Menu tmPURCHmonth_year 
            Caption         =   "„ﬁ«—‰… „‘ —Ì«  «·‘ÂÊ— ··”‰Ê« "
         End
         Begin VB.Menu tmPURCHmonth_MODEL 
            Caption         =   "„ﬁ«—‰… „‘ —Ì«  «·„ÊœÌ·«  ··”‰Ê« "
         End
      End
      Begin VB.Menu tmrepsal 
         Caption         =   " ﬁ«—Ì— «·„»Ì⁄«  «·›—Ê⁄ Ê «· ÊﬂÌ·« --"
         Begin VB.Menu tmsalesrepall 
            Caption         =   " ﬁ«—Ì— „»Ì⁄«  „Ã„⁄"
         End
         Begin VB.Menu tmbalonest 
            Caption         =   "„»Ì⁄«  & —’Ìœ „ÊœÌ·«  «·«„«‰« - ›Ê« Ì— «„«‰« "
         End
         Begin VB.Menu tmmansales 
            Caption         =   "≈Ã„«·Ï „»Ì⁄«  «ﬁ”«„ »«∆⁄Ì‰ ⁄œœ & ﬁÌ„…"
         End
         Begin VB.Menu tmsalavgman 
            Caption         =   "„»Ì⁄«  »«∆⁄Ì‰ - ≈Ã„«·Ï & „ Ê”ÿ „»Ì⁄« "
         End
         Begin VB.Menu tmproftbramonth 
            Caption         =   "≈Ã„«·Ï «—»«Õ ‘Â—Ì…  "
         End
         Begin VB.Menu tm_grdProfitMonth 
            Caption         =   "≈Ã„«·Ì «—»«Õ ‘Â—Ì… ··›—Ê⁄"
         End
         Begin VB.Menu tmsalesmodel 
            Caption         =   "„»Ì⁄«  „ÊœÌ·«  ⁄œœ - ”⁄— „’‰⁄- ﬁÌ„… "
         End
         Begin VB.Menu tmsalesokaz 
            Caption         =   "≈Ã„«·Ï „»Ì⁄«  & Œ’„ «Êﬂ«“ÊÌ‰ & ⁄—Ê÷ "
         End
         Begin VB.Menu tmsalesokaz_br 
            Caption         =   "≈Ã„«·Ï „»Ì⁄«  & Œ’„ «Êﬂ«“ÊÌ‰ & ⁄—Ê÷  - ··›—Ê⁄"
         End
         Begin VB.Menu tmRepSales12 
            Caption         =   " ›’Ì·Ï ﬁÌ„… „»Ì⁄«  „ÊœÌ·«   »”⁄— „” Â·ﬂ & »”⁄— «Êﬂ«“ÊÌ‰"
         End
         Begin VB.Menu tmsales_barcode 
            Caption         =   " ›’Ì·Ï ﬂ„Ì«  „»Ì⁄«  «·Ê«‰ - „ﬁ«”«  - »«—ﬂÊœ ··›—Ê⁄"
         End
         Begin VB.Menu tmsalesdiscdoc 
            Caption         =   "„»Ì⁄«  „ÊœÌ·«  «·⁄—Ê÷ ··„ÊœÌ·« "
         End
         Begin VB.Menu tmsalespricedisctype 
            Caption         =   "„»Ì⁄«  »”⁄— „” Â·ﬂ - ⁄—Ê÷ ··›—Ê⁄"
         End
         Begin VB.Menu tmsalesretbr 
            Caption         =   "≈Ã„«·Ï „»Ì⁄«  - „— Ã⁄ - „— Ã⁄ „‰ ›—Ê⁄ «Œ—Ï"
         End
      End
      Begin VB.Menu tmsalesavg 
         Caption         =   " ﬁ«—Ì— „ Ê”ÿ «·„»Ì⁄« --"
         Begin VB.Menu tmsalesavgman 
            Caption         =   "„»Ì⁄«  »«∆⁄Ì‰ - ≈Ã„«·Ï & „ Ê”ÿ „»Ì⁄« "
         End
         Begin VB.Menu tmsalesavgbranch 
            Caption         =   "„»Ì⁄«  Ê „ Ê”ÿ „»Ì⁄«  «·›—Ê⁄"
         End
         Begin VB.Menu tmsalesavgsec 
            Caption         =   "„»Ì⁄«  Ê „ Ê”ÿ „»Ì⁄«  «·«ﬁ”«„"
         End
         Begin VB.Menu tmsalesavgsupler 
            Caption         =   "„»Ì⁄«  Ê „ Ê”ÿ „»Ì⁄«  «·„Ê—œÌ‰"
         End
      End
      Begin VB.Menu tmpurchrep 
         Caption         =   " ﬁ«—Ì— „‘ —Ì«  „ÊœÌ·« --"
         Begin VB.Menu tmpurchmodel 
            Caption         =   "≈Ã„«·Ï „‘ —Ì«  „ÊœÌ·« "
         End
         Begin VB.Menu tmpurchSEC 
            Caption         =   "≈Ã„«·Ï „‘ —Ì«  «ﬁ”«„"
         End
         Begin VB.Menu tmpurchsupler 
            Caption         =   "≈Ã„«·Ï „‘ —Ì«  „Ê—œÌ‰"
         End
         Begin VB.Menu tmpurchstore 
            Caption         =   "≈Ã„«·Ï „‘ —Ì«  «·›—Ê⁄"
         End
         Begin VB.Menu tmpurchdoc 
            Caption         =   "≈Ã„«·Ï „” ‰œ«  „‘ —Ì« "
         End
      End
      Begin VB.Menu tmbranchshop 
         Caption         =   " ﬁ«—Ì— „ﬁ«—‰… «·›—Ê⁄ Ê «· ÊﬂÌ·« --"
         Begin VB.Menu tmrepbr5 
            Caption         =   " Ê“Ì⁄ ⁄œœ ﬁÌ„… —’Ìœ »÷«⁄… „Ê—œÌ‰ ··›—Ê⁄"
         End
         Begin VB.Menu tmrepbr1 
            Caption         =   " Ê“Ì⁄ „»Ì⁄«  „ÊœÌ·«  ··›—Ê⁄"
         End
         Begin VB.Menu tmrepbr2 
            Caption         =   " Ê“Ì⁄ „»Ì⁄«  «·„Ê—œÌ‰ ··›—Ê⁄"
         End
         Begin VB.Menu tmrepbr3 
            Caption         =   " Ê“Ì⁄ „»Ì⁄«  ÌÊ„Ì… ··›—Ê⁄"
         End
         Begin VB.Menu tmsalesmonthbr 
            Caption         =   " Ê“Ì⁄ „»Ì⁄«  ‘Â—Ì… ··›—Ê⁄"
         End
         Begin VB.Menu tmsecbransh 
            Caption         =   " Ê“Ì⁄ „»Ì⁄«  «·«ﬁ”«„ ··›—Ê⁄"
         End
         Begin VB.Menu tmgrbranch 
            Caption         =   " Ê“Ì⁄ „»Ì⁄«  «·„Ã„Ê⁄«  ··›—Ê⁄"
         End
         Begin VB.Menu tmallbranch 
            Caption         =   "≈Ã„«·Ï „Êﬁ› «·›—Ê⁄"
         End
         Begin VB.Menu tmproftsales 
            Caption         =   "≈Ã„«·Ï «—»«Õ «·›—Ê⁄"
         End
         Begin VB.Menu tm_grd_store_profit 
            Caption         =   "≈Ã„«·Ì «—»«Õ «·›—Ê⁄ ⁄·Ì «·‘ÂÊ—"
         End
      End
      Begin VB.Menu tmbalcust 
         Caption         =   "«—’œ… «·⁄„·«¡"
      End
      Begin VB.Menu tmrepsupl 
         Caption         =   " ﬁ—Ì— «·„Ê—œÌ‰--"
         Begin VB.Menu tmbalsupl 
            Caption         =   "√—’œ… Ê  ⁄«„·«  «·„Ê—œÌ‰ - „Ê«”„"
         End
         Begin VB.Menu tmpaysupp 
            Caption         =   "œ›⁄«  „” Õﬁ… ··„Ê—œÌ‰"
         End
         Begin VB.Menu tmproftsupler 
            Caption         =   "«—»«Õ «·„Ê—œÌ‰"
         End
         Begin VB.Menu tmbalsupl2 
            Caption         =   "√—’œ… Ê  ⁄«„·«  «·„Ê—œÌ‰"
         End
      End
      Begin VB.Menu tmbalscal 
         Caption         =   " ﬁ—Ì— Ã—œ Ê  ﬁÌ„ «·„ÊœÌ·« "
      End
      Begin VB.Menu storerep 
         Caption         =   " ﬁ«—Ì— «’‰«›  & „»Ì⁄« "
      End
      Begin VB.Menu tmdocinout 
         Caption         =   "≈Ã„«·Ï „” ‰œ«  ’«œ— ÊÊ«—œ ··›—Ê⁄"
      End
      Begin VB.Menu tmmodeltrans 
         Caption         =   " ›’Ì·Ï Ê«—œ Ê ’«œ— „ÊœÌ·«  «·„ÊœÌ·«  ··›—Ê⁄"
      End
      Begin VB.Menu m_SupRpt 
         Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰ 2"
      End
      Begin VB.Menu tmRepCash 
         Caption         =   " ﬁ«—Ì— „’«—Ì› & ‰ﬁœÌ…--"
         Begin VB.Menu tmChargerep 
            Caption         =   " ﬁ«—Ì—  „’«—Ì› & ≈Ì—«œ«  & ‘—ﬂ«¡"
         End
         Begin VB.Menu tmchargetear1 
            Caption         =   "„ﬁ«—‰… „’«—Ì› ›—Ê⁄ ··”‰Ê« "
         End
         Begin VB.Menu tmchargetear2 
            Caption         =   "≈Ã„«·Ï  „’«—Ì› «·›—Ê⁄ - „’«—Ì› —∆Ì”Ì…"
         End
         Begin VB.Menu tmrepchagebr 
            Caption         =   "„’«—Ì› ›—Ê⁄  ÊﬂÌ·« "
         End
      End
      Begin VB.Menu tmChqRep 
         Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
      End
      Begin VB.Menu tmproftrep 
         Caption         =   "ﬁ«∆„… «·œŒ·"
      End
      Begin VB.Menu tmbaltot 
         Caption         =   "≈Ã„«·Ï «—’œ…"
      End
   End
   Begin VB.Menu mnServices 
      Caption         =   "Œœ„« "
      Begin VB.Menu tmsecurity 
         Caption         =   "’·«ÕÌ« "
      End
      Begin VB.Menu tm_logo_codes 
         Caption         =   "»Ì«‰«  «·‘⁄«—« "
      End
      Begin VB.Menu tmsubuser 
         Caption         =   "’·«ÕÌ«  „— Ã⁄ Ê Œ’„"
      End
      Begin VB.Menu tmzone 
         Caption         =   "»Ì«‰«  „‰«ÿﬁ ›—Ê⁄"
      End
      Begin VB.Menu tmbranchdata 
         Caption         =   "»Ì«‰«  «·›—Ê⁄"
      End
      Begin VB.Menu tmbranch 
         Caption         =   " ”⁄Ì— «Êﬂ«“ÌÊ‰ ··›—Ê⁄"
      End
      Begin VB.Menu tmvisa 
         Caption         =   "»Ì«‰«  «·›Ì“«"
      End
      Begin VB.Menu tmaddres 
         Caption         =   "»Ì«‰«  «·›—⁄"
      End
      Begin VB.Menu xMSal 
         Caption         =   "√ﬂÊ«œ »«∆⁄Ì‰"
      End
      Begin VB.Menu tm_printers 
         Caption         =   "÷»ÿ «·ÿ«»⁄« "
      End
      Begin VB.Menu tmpriceokaz 
         Caption         =   " ÕœÌœ ”⁄— «·»Ì⁄"
      End
      Begin VB.Menu tmdiscshop 
         Caption         =   "Œ’Ê„«  «·„Õ·«  ·› —…"
      End
      Begin VB.Menu tmmosm 
         Caption         =   "»Ì«‰«  «·„Ê«”„"
      End
      Begin VB.Menu tmshipcode 
         Caption         =   "‘—ﬂ«  «·‘Õ‰"
      End
      Begin VB.Menu tmsenddataip 
         Caption         =   " ÕœÌÀ «·›—Ê⁄  IP"
      End
      Begin VB.Menu tmrebranch 
         Caption         =   " »œÌ· »‰ «·›—Ê⁄"
      End
      Begin VB.Menu tmnewmosm 
         Caption         =   "› Õ „Ê”„ ÃœÌœ"
      End
      Begin VB.Menu tmattac 
         Caption         =   "”Õ» »Ì«‰«  «·Õ÷Ê—"
      End
      Begin VB.Menu tmBackUp 
         Caption         =   "BACKUP DATA"
      End
      Begin VB.Menu tm_einvoice 
         Caption         =   "«⁄œ«œ«  «·›« Ê—… «·«·Ìﬂ —Ê‰Ì…"
      End
      Begin VB.Menu tm_send_invoices 
         Caption         =   "√—”«· ›Ê« Ì— «·›—⁄"
      End
      Begin VB.Menu xExit 
         Caption         =   "Œ—ÊÃ"
      End
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub MainMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.key)
Case "XSALES"
    salesfrm.myPublic = 0
    salesfrm.Show 1
Case "XBDATA"
    itemsgrdfrm.Show
Case "XEXIT"
    Unload Me
End Select
End Sub
Private Sub MDIForm_Load()
Dim TableAddress As New ADODB.RecordSet

'SetKbLayout Lang_EN
If cusername <> "Aymen Junior" Then
    If Not bSupermode Then LoadMenu
Else
    mnitems.Visible = False
    mnclients.Visible = False
    
    
    mnVendorsx.Visible = False
    mnInvoice.Visible = False
    mnMAINCash.Visible = False
    mnBank.Visible = False
    mnCASH.Visible = False
    mnServices.Visible = True
    
    
    tmsubuser.Visible = False
    tmzone.Visible = False
    tmbranchdata.Visible = False
    tmbranch.Visible = False
    tmbranchBR.Visible = False
    tmvisa.Visible = False
    xMSal.Visible = False
    tmpriceokaz.Visible = False
    tmdiscshop.Visible = False
    tmmosm.Visible = False
    tmshipcode.Visible = False
    tmsenddataip.Visible = False
    tmsenddataip_fr.Visible = False
    tmrebranch.Visible = False
    tmnewmosm.Visible = False
    tmattac.Visible = False
    tm_einvoice.Visible = False
    tm_send_invoices.Visible = False
    tmmodelfact.Visible = False
End If

If con.State = adStateOpen Then con.Close
OpenCon con


'''''''  ADDRESS
TableAddress.Open "ADDRESS", con, adOpenStatic, adLockReadOnly, adCmdTable
If TableAddress.RecordCount > 0 Then
    cComp_Name = TableAddress!DESCA
End If

If lSupperVisor Or bOpt5 Or cBranch <> "00" Then
    cUserStore = " CODE IS NOT NULL ORDER BY CODE "
Else
    cUserStore = " CODE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & MyParn(nusercode) & " ) " & "  ORDER BY CODE "
End If
If cComp_Name = "" Then cComp_Name = "JUNIOR"
Firsttitle = cComp_Name
If cBranch <> "00" Then
    bOpt5 = True
    pServerIp = GetDesca("SELECT IPSERVER FROM BRANCH", con)
    If pServerIp = "ONLINE" Then lServerOnLineShop = True
    pServerData = GetDesca("SELECT DATASERVER FROM BRANCH", con)
    If lServerOnLineShop Then
        'pServerIp = "junior-sql.database.windows.net"
    End If
    pServerData = GetDesca("SELECT DATASERVER FROM BRANCH", con)
End If
If cPMosm = "" Then cPMosm = GetDesca("select mosm from mosm order by date desc ", con)

CheckData
'
ClosedCashDoc

Dim PriceTable As New ADODB.RecordSet
PriceTable.Open "PRICE", con, adOpenStatic, adLockReadOnly, adCmdTable

Dim File0_00Table As New ADODB.RecordSet
File0_00Table.Open "FILE0_00", con, adOpenStatic, adLockReadOnly, adCmdTable

If PriceTable.RecordCount > 0 Then
    cRet = PriceTable!price
    lokaz = IIf(cRet = "True", True, False)
    
    'If File0_00Table.RecordCount > 0 Then cRet = File0_00Table!OKAZ_ITEM2
    lokazItem = IIf(cRet = "True", True, False)
    lOneVisa = IIf(Val(GetDesca("SELECT COUNT(CODE) FROM VISA ", con) & "") = 1, True, False)
    If lOneVisa Then sCodeVisaBranch = GetDesca("SELECT CODE FROM VISA ", con)
End If

If cBranch = "00" Then
    dSalesDate = Date
Else
    If lIsBranchStore Then
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES WHERE BRANCH  = " & MyParn(cBranch), con), "DD-MM-YYYY")
        If Not IsDate(dSalesDate) Then con.Execute " INSERT INTO [dsales] (BRANCH,dsales) VALUES ( " & addstring(cBranch) & " ," & addDate(Date) & " ) "
        dSalesDate = Date
    Else
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES", con), "DD-MM-YYYY")
    End If
End If
If IsDate(dSalesDate) Then
    dMaxDate = DateAdd("D", -4, dSalesDate)
Else
    dMaxDate = Date
    dSalesDate = Date
End If

If Not IsDate(dSalesDate) Then dSalesDate = DateValue("01-01-2000")

Me.Caption = Firsttitle

sStore = GetDesca("SELECT STORE FROM BRANCH WHERE CODE = " & MyParn(cBranch), con)
cBranchStore = sStore
If cUserBox = "" Then cBranchBox = ""

On Error Resume Next
tmrebranch.Visible = lShowBranch
If sBranchCode <> "00" Then
    tm_einvoice.Visible = False
    tmrepinvoice.Visible = False
    tmorderamazon.Visible = False
    tmRecevdOrder.Visible = False
    tmorder_rec_amazon.Visible = False
    tmrecordercust.Visible = False
    tmrecorder_supler.Visible = False
    tmOrderDem.Visible = False
    tmOrderDem_ALL.Visible = False
    tmOrderDOC_ALL.Visible = False
    
    tmrepinvoice.Visible = False
    tmzone.Visible = False
    tmsalesinvret.Visible = False
    tmTargetBr.Visible = False
    tmsalestax.Visible = False
    tmSalesBranch.Visible = False
    tmbalbrach.Visible = False
    
    If tmSalesBranch.Visible Then tmSalesBranch.Visible = Not lIsBranchStore
'   If tmbalbrach.Visible Then tmbalbrach.Visible = Not lIsBranchStore
    If xCASH2.Visible Then xCASH2.Visible = Not lIsBranchStore
    If xmcash3.Visible Then xmcash3.Visible = Not lIsBranchStore
    If tmchargecode.Visible Then tmchargecode.Visible = Not lIsBranchStore
    If tmchargemaincode.Visible Then tmchargemaincode.Visible = Not lIsBranchStore
    If tmcashbranch.Visible Then tmcashbranch.Visible = Not lIsBranchStore
    tmvisa.Visible = False
'    tmransbrancip.Visible = False
    tmsakesstore.Visible = False
'    If tmNewDoc.Visible Then tmNewDoc.Visible = Not lIsBranchStore
    tmbaltot.Visible = False
    If tmproftrep.Visible Then tmproftrep.Visible = (cBranch > "60")
'    tmChargerep.Visible = False
    tmbalbox.Visible = False
    tmNewDocBranch.Visible = False
    
    tmbox.Visible = False
    tmpart_code.Visible = False
    tmmansales2.Visible = False
    tmmansales2.Visible = False
'    tmsalday.Visible = False
    tmOrderDem.Visible = False
    tmpurchordertot.Visible = False
    tmorderpurch.Visible = False
    tmrepdemorder.Visible = False
    tmsaldayfr.Visible = False
    tmDashBoard.Visible = False
    tmbranchBR.Visible = False
'    tmfixcostsales.Visible = False
    tmrepalltrans.Visible = False
    tmrepalltrans2.Visible = False
    tmsubuser.Visible = False
    tmdataitem.Visible = False
    tmdocsend.Visible = False
    tmpirchokbr.Visible = False
    tmtrans_br.Visible = False
    xItemMoveFR.Visible = False
    If Not bSupermode Then
        tmmodelokazgr.Visible = False
    End If
    tmgrmain.Visible = False
    tmsexcode.Visible = False
    tmagecode.Visible = False
    tmdisccardprint.Visible = False
    If sBranchCode > "60" Then
'        If tmransbrancip.Visible Then tmransbrancip.Visible = True
        
        If tmcash0.Visible Then tmcash0.Visible = True
        If xmcash3.Visible Then xmcash3.Visible = True
        If xCashed.Visible Then xmcash3.Visible = True
        If tmcash11.Visible Then tmcash11.Visible = True
        
        If tmrepsupl.Visible Then tmrepsupl.Visible = True
        If tmincomemaincode.Visible Then tmincomemaincode.Visible = True
        If tmincomecode.Visible Then tmincomecode.Visible = True
        If xItemMove.Visible Then xItemMove.Visible = True
        If tm_vsitem.Visible Then tm_vsitem.Visible = True
        If tmfixcostsales.Visible Then tmfixcostsales.Visible = True
        If mnclients.Visible Then mnclients.Visible = True
        If mnVendorsx.Visible Then mnVendorsx.Visible = True
        If mnBank.Visible Then mnBank.Visible = True
'       tmtrans.Visible = False
        If tmmodels.Visible Then tmmodels.Visible = True
        tmbalbrach.Visible = False
        If tmrepalltrans.Visible Then tmrepalltrans.Visible = True
        If tmrepalltrans2.Visible Then tmrepalltrans2.Visible = True
        If tmrepsupl.Visible Then tmrepsupl.Visible = True
        tmdocinout.Visible = False
        tmmodeltrans.Visible = False
        tmtranstocust.Visible = False
        tmTRANSFROMIP_FR.Visible = False
        If xMSal.Visible Then xMSal.Visible = True
        If tmmosmsecbal.Visible Then tmmosmsecbal.Visible = True
    Else
        xmcash3.Visible = False
        tmcash0.Visible = False
        xCashed.Visible = False
        tmcash11.Visible = False
        
'        tmransbrancip.Visible = False
        tmrepsupl.Visible = False
        tmincomemaincode.Visible = False
        tmincomecode.Visible = False
        tmshipcode.Visible = False
        tm_vsitem.Visible = False
        xItemMoveFR.Visible = False
        tmmosmsecbal.Visible = False
        xMSal.Visible = False
        tmTRANSFROMIP_FR.Visible = False
        m_SupRpt.Visible = False
        tmbalbrachFR.Visible = False
'       tm_vsitem.Visible = False
        tmmodels.Visible = False
        mnclients.Visible = False
        mnVendorsx.Visible = False
        mnBank.Visible = False
        tmnewmosm.Visible = False
    End If
    trmtrans3.Visible = False
    tmsenddataip_fr.Visible = False
'    tmtransdata.Visible = False
    tmpriceokaz.Visible = False
    tmpurchbrret.Visible = False
    tmpurchbr.Visible = False
    
    
    tmtranscost.Visible = False
    
    
    tmitemgroup.Visible = False
    tmstores.Visible = False
    tmfact.Visible = False
    tmsec.Visible = False
    
    tmbalstore.Visible = False
    sep101.Visible = False
    tmmosmgrbal.Visible = False
    tmmosm.Visible = False
    sep65.Visible = False
    sep70.Visible = False
    sep69.Visible = False
    tmrepeicefact.Visible = False
    tmbranchshop.Visible = False
    tmsenddataip.Visible = False
    tmtranstocust.Visible = False
    tmreorder.Visible = False
    tmsalesmodel.Visible = False
    tmChqRep.Visible = False
    tmbranch.Visible = False
    tmrepbranchfr.Visible = False
    tmdiscshop.Visible = False
    tmallbranchsals.Visible = False
    storerep.Visible = False
    sep5_1.Visible = False
    xCashed.Visible = False
    xmcash4.Visible = False
    tmtranscode2.Visible = False
    If tmtransbarcode.Visible Then tmtransbarcode.Visible = True
Else
    tmtransbarcode.Visible = False
'    tmransbrancip.Visible = True
    trmtrans2.Caption = " ÕÊÌ· Ê«—œ „‰ «·›—Ê⁄"
    trmtrans3.Caption = " ÕÊÌ· Ê«—œ „‰ «· ÊﬂÌ·« "
End If

If cusername = "Aymen Junior" Then
    tmsecurity.Visible = True
    tm_printers.Visible = True
    tmaddres.Visible = True
End If

tmbranchdata.Visible = bSupermode And cBranch = "00"

nCountBranch = Val(GetDesca("SELECT MAX(CODE ) FROM BRANCH ", con) & "")
nCountBranch_fr = Val(GetDesca("SELECT MAX(CODE ) FROM BRANCH_FR ", con) & "")

If Not lServerOnLine Then
    'checkCopy
End If
If cBranch <> "00" And Not bopt2 And Not lShowBranch Then
    salesfrm.Show
End If
If retFlag(aBranch, "code") = 1 Then Me.BackColor = RGB(226, 218, 209)
Err.Clear
End Sub
Private Sub LoadMenu()
Dim cConStr As String
If lShowBranch Then
    cConStr = LoadConStringServer
    OpenCon con, cConStr
Else
    OpenCon con
End If
Dim sectable As New ADODB.RecordSet
cString = "SELECT users.code,users.title1,users.title2, users.Password, users.desca,USERS.option1,users.option2,users.option3,users.option4,users.option5 ,Menusetting.Visible, Menusetting.Editable, Menu.Control, Menu.MenuNo,Menu.mainMenu " & _
          " FROM (users INNER JOIN Menusetting ON users.Code = Menusetting.code) INNER JOIN Menu ON Menusetting.control = Menu.Control " & _
          " where users.code = " & nusercode
sectable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
For i = 0 To Main.Count - 1
    If TypeOf Main(i) Is Menu And Mid(Main(i).Name, 1, 2) = "mn" Then
        Main(i).Visible = False
    End If
Next
Err.Clear
On Error Resume Next
For i = 0 To Main.Count - 1
    If "tmonline_order" = Main(i).Name Then
        A = A
    End If
    If TypeOf Main(i) Is Menu And Mid(Main(i).Name, 1, 2) <> "mn" Then
        sectable.Find "control = " & MyParn(Main(i).Name), , adSearchForward, adBookmarkFirst
        If Not sectable.EOF Then
            If sectable!Visible Then
                Main(sectable!MainMenu).Visible = True
                Main(i).Visible = True
            Else
'                For i2 = 1 To MainMenu.Buttons.Count
'                    If UCase(MainMenu.Buttons(i2).key) = UCase(Main(i).Name) Then
'                        MainMenu.Buttons(i2).Enabled = False
'                    End If
'                Next
                Main(i).Visible = False
            End If
        Else
            Main(i).Visible = False
        End If
    End If
Next
Err.Clear
End Sub
Private Sub Stores_Click()
End Sub
Private Sub report1_Click()
Load Form4
Form4.Show 1
End Sub
Private Sub bal_box_Click()
    Load BalBox
    BalBox.Show 1
End Sub
Private Sub BANK_MOVE_Click()
Vs_Bank.Show 1
End Sub
Private Sub CHQREP_Click()
Load repchq
repchq.Show 1
End Sub
Private Sub m_ChargeMain_Click()
Flag.myFlag = 7
Flag.Show 1
End Sub
Private Sub m_debtCode_Click()
publicFlag = 1
Load AssetCode
AssetCode.Show 1
End Sub
Private Sub m_DebtGrp_Click()
publicFlag = 15
Load Flag
Flag.Show 1
End Sub
Private Sub m_Income_Click()
publicFlag = 1
Load ChargeSub
ChargeSub.Show 1
End Sub

Private Sub m_IncomeMain_Click()
publicFlag = 13
Load Flag
Flag.Show 1
End Sub
Private Sub M_SALMED_Click()
Load SalMed
SalMed.Show 1
End Sub

Private Sub itChargImp_Click()
ReDim aPublic(5)
aPublic(0) = "FILE7_60CH_CODE"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„’—Ê›"
aPublic(4) = "»Ì«‰ "
aPublic(5) = "√ﬂÊ«œ „’«—Ì› ≈” Ì—«œÌ…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1

End Sub

Private Sub ititemimport_Click()
    VsImpItem.Show
End Sub

Private Sub ItPostItem_Click()
    CopyItem.Show 1
End Sub
Private Sub m_SupRpt_Click()
    rpSup.Show
End Sub
Private Sub mChargerep_Click()
    RepCharge.Show 1
End Sub
Private Sub MCURR_Click()
publicFlag = 6
Load Flag
Flag.Show 1
End Sub
Private Sub mInput_Click()
publicFlag = 0
Load Vs_Input
Vs_Input.Show 1
End Sub

Private Sub MnTSalMan_Click()
vsman.Show 1
End Sub

Private Sub Mony_Box_Click()
publicFlag = 5
Flag.Show 1
End Sub
Private Sub mPricelist_Click()
Load PriceList
PriceList.Show 1
End Sub
Private Sub mTransAction_Click()
Load Post
Post.Show 1
End Sub
Private Sub mPrevInv_Click()
    Load PrevInv
    PrevInv.Show 1
End Sub

Private Sub mxclose_Click()
FrmClose.Show 1
End Sub

Private Sub mXsec_Click()
    Security.Show 1
End Sub

Private Sub mxtrans_Click()
    TransBox.Show 1
End Sub
Private Sub PHO_MAI_Click()
Load phone
phone.Show 1
End Sub
Private Sub RBAN_Click()
    Load RepBank
    RepBank.Show 1
End Sub
Private Sub utillll_Click()
    Load Utility2
    Utility2.Show 1
End Sub
Private Sub Trans_Box_Click()
    TransBox.Show 1
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub tm_complain_Click()
End Sub

Private Sub tm_sup_orders_Click()
    
End Sub

Private Sub TEST_Click()
Form1.Show
End Sub

Private Sub TEST2_Click()
    TestDisplay.Show 1
End Sub

Private Sub rmdocstock_Click()
    VsTStockDoc.Show

End Sub

Private Sub tm_box_day_Click()
BoxDayfrm.Show
End Sub

Private Sub tm_einvoice_Click()
einvoice_menufrm.Show 1
End Sub

Private Sub tm_grd_input_Click()
grdMonthInput.Show
End Sub

Private Sub tm_grd_store_profit_Click()
grdStoreProfit.Show
End Sub

Private Sub tm_grdBalancePrice_Click()
grdBalancePrice.Show
End Sub

Private Sub tm_grdProfitMonth_Click()
grdMonthProfit.Show
End Sub

Private Sub tm_grid_inv_balance_Click()
grdInvBalance.Show
End Sub

Private Sub tm_list_price_Click()
price_listfrm.Show
End Sub

Private Sub tm_logo_codes_Click()
logofrm.Show 1
End Sub

Private Sub tm_online_orders_Click()
Online_Stage_main.Show
End Sub

Private Sub tm_printers_Click()
    printersfrm.Show 1
End Sub

Private Sub tm_sup_okaz_Click()
    okazfrm.Show
End Sub
 
Private Sub tm_sales_orders_Click()
sales_ordersfrm.Show
End Sub

Private Sub tm_send_invoices_Click()
    InvoiceTransfrm.Show 1
End Sub

Private Sub tm_vsitem_Click()
    If cBranch = "00" Then
        VsTItem.Show
    Else
        VsTitem_FR.Show
    End If
End Sub

Private Sub tmacccode_Click()
    AccCode.Show
End Sub

Private Sub tmaddres_Click()
'    If cBranch <> "00" Then
    SettingFrm.Show 1
End Sub

Private Sub tmagecode_Click()
    ReDim aLocal(6)
    aLocal(0) = "FILE0_82"
    aLocal(1) = "Code"
    aLocal(2) = "Desca"
    aLocal(3) = "ﬂÊœ «·›∆…"
    aLocal(4) = "»Ì«‰ "
    aLocal(5) = " ”ÃÌ· «·›∆«  «·⁄„—Ì…"
    aLocal(6) = 2
    FlagFrm2.bEdit = True
    FlagFrm2.myPublic = aLocal
    FlagFrm2.Show 1
End Sub

Private Sub tmallbranch_Click()
    grditemfrm1.Show
End Sub

Private Sub tmallbranchsals_Click()
    VsTfrBranshSalModel_ALL.Show
End Sub
Private Sub tmallSEC_Click()
    VsTSection.Show
End Sub
Private Sub tmallsupp_Click()
    VsTSupler.Show
End Sub

Private Sub tmattac_Click()
    machineGetFrm.Show 1
End Sub

Private Sub tmBackUp_Click()
    copyFlashfrm.Show 1
End Sub

Private Sub tmbalacc_Click()
    VsTBalAcc.Show
End Sub

Private Sub tmbalage_Click()
    VsTbalage.Show
End Sub

Private Sub tmbalbox_Click()
    VsTCashDay.Show
End Sub
Private Sub tmbalbrach_Click()
    grditem_br.Show
End Sub

Private Sub tmbalbrachFR_Click()
    grditem_br_FR.Show
End Sub

Private Sub tmbalcust_Click()
    'VsTBalCust.Show
    grdcust_balancefrm.Show
End Sub

Private Sub tmbalmodecode_Click()
    TPrice_Supl.Show
End Sub

Private Sub tmbalmodelgr_Click()
    TPrice_group.Show
End Sub

Private Sub tmbalmodelscal_Click()
    VsTbal_Scal.Show
End Sub

Private Sub tmbalmodelsec_Click()
    TPrice_SEC.Show
End Sub

Private Sub tmbalmodemodel_Click()
    TPrice_item.Show
End Sub

Private Sub tmbalonest_Click()
    VsTAllOnest.Show
End Sub

Private Sub tmbalscal_Click()
    rpitem4.Show 1
End Sub

Private Sub tmbalshop_Click()
'    grditemfrm3.Show
End Sub

Private Sub tmbalstore_Click()
    VsTStore.Show
End Sub

Private Sub tmbaLSTORE_PRICE_Click()
    TPrice_STORE.Show
End Sub

Private Sub tmbalstore_scal_Click()
    VsBalItem_Store.Show
End Sub

Private Sub tmbalsupl_Click()
    VsTBalSupl.Show
End Sub

Private Sub tmbalsupl2_Click()
   'VsTBalSUPL_ALL.Show
    grdSupBalance.Show
End Sub
Private Sub tmbaltot_Click()
    frmTot.Show
End Sub

Private Sub tmBankData_Click()
    bankfrm.bEdit = True
    bankfrm.Show
End Sub
Private Sub tmbankDtl1_Click()
    rpBank3.Show
End Sub
Private Sub tmBankGrroup_Click()
ReDim aPublic(5)
aPublic(0) = "FILE5_50"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„Ã„Ê⁄…"
aPublic(4) = "≈”„ «·„Ã„Ê⁄…"
aPublic(5) = "„Ã„Ê⁄«  «·»‰Êﬂ"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub
Private Sub tmBankInout_Click()
    BankInOutfrm.bEdit = True
    BankInOutfrm.Show
End Sub
Private Sub tmBankItems_Click()
ReDim aPublic(5)
aPublic(0) = "FILE5_00"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·Õ—ﬂ…"
aPublic(4) = "≈”„ «·Õ—ﬂ…"
aPublic(5) = "«ﬂÊ«œ Õ—ﬂ… «·»‰Êﬂ"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show
End Sub
Private Sub tmBankMove_Click()
    BankMovefrm.Show
End Sub
Private Sub tmBankMoveTotal_Click()
    rpBank2.Show
End Sub
Private Sub tmBankState_Click()
    bankStatefrm.Show
End Sub

Private Sub tmbarcodeprint_Click()
    Dream_Bar.Show
End Sub
Private Sub tmbarcode_Click()
'If retFlag(aBranch, "code") = "1" Then
'    barcodefrm2.Show 1
'Else
    BarCodeFRM.Show 1
'End If
End Sub

Private Sub tmbarcode2_Click()
    barcodefrm_2.Show
End Sub
Private Sub tmbox_Click()
    Boxfrm.bEdit = True
    Boxfrm.Show
End Sub
Private Sub tmboxbal_Click()
    BalBox.Show
End Sub
Private Sub tmboxtrans_Click()
    If lIsBranchStore Then
        boxtrans_2.bEdit = True
        boxtrans_2.Show
    Else
        boxtransfrm.bEdit = True
        boxtransfrm.Show
    End If
End Sub

Private Sub tmbranch_Click()
    branch_codesfrm.Show
End Sub

Private Sub tmbranchBR_Click()
    StoreData_BR.Show
'   branchbr_codesfrm.Show
End Sub

Private Sub tmbranchdata_Click()
    branch.Show
End Sub

Private Sub tmcash0_Click()
    Cashfrm.myPublic = 0
    Cashfrm.Show
    End Sub

Private Sub tmcash11_Click()
    Cashfrm.myPublic = 11
    Cashfrm.Show
End Sub

Private Sub tmcashbranch_Click()
    chargefrm.bEdit = True
    chargefrm.myPublic = 3
    chargefrm.Show
End Sub

Private Sub tmcashser_Click()
    rpSup11.Show 1
End Sub

Private Sub tmcharge_Click()
    chargefrm.bEdit = True
    chargefrm.myPublic = 1
    chargefrm.Show
End Sub

Private Sub tmcharge2_Click()
    chargefrm.bEdit = True
    chargefrm.myPublic = 4
    chargefrm.Show
End Sub
Private Sub tmchargecode_Click()
'    chargecodefrm.bEdit = True
'    chargecodefrm.myPublic = 1
'    chargecodefrm.Show 1
chargeCodesfrm.bEdit = True
chargeCodesfrm.myPublic = 1
chargeCodesfrm.Show 1
End Sub
Private Sub tmchargemaincode_Click()
ReDim aPublic(6)
aPublic(0) = "FILE8_52"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·„’—Ê›"
aPublic(4) = "»Ì«‰ «·„’—Ê›"
aPublic(5) = "„’«—Ì› —∆Ì”Ì…"
aPublic(6) = 3
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aPublic
FlagFrm2.Show 1
End Sub
Private Sub tmChargerep_Click()
    rpCharge.Show
End Sub

Private Sub tmchargetear1_Click()
    VsTCharge_YEAR.Show
End Sub

Private Sub tmchargetear2_Click()
    VsTCharge_Branch.Show
End Sub

Private Sub tmChqIn_Click()

    publicFlag = 1
    bEdit = True
    chqClientfrm.Show 1
End Sub
Private Sub tmChqOut_Click()
publicFlag = 2
bEdit = True
chqsupfrm.Show 1
End Sub
Private Sub tmChqRep_Click()
    rpChq.Show
End Sub
Private Sub tmClientReport_Click()
    rpClient.Show
End Sub

Private Sub tmcolor_Click()
colorfrm.bEdit = True
colorfrm.Show
End Sub
Private Sub TMCustImp_Click()
CustSalesImp.Show
End Sub
Private Sub tmCut_Click()
Cutfrm.bEdit = True
Cutfrm.Show
End Sub
Private Sub tmDamage_Click()
DamageFrm.myPublic = 1
DamageFrm.Show
End Sub
Private Sub tmchqser_Click()
    rpSup22.Show 1
End Sub

Private Sub tmcomp_Click()
    TComp1.Show 1
End Sub

Private Sub tmcopymdb_Click()
    CopyData.Show
End Sub

Private Sub tmcostbalbransh_Click()
    VsTFact.Show
End Sub
Private Sub tmcust_Click()
    Clients.Show
End Sub
Private Sub tmcustcash_Click()
    Clients_sub.Show 1
End Sub

Private Sub tmDashBoard_Click()
    TSal_DashBoard.Show
End Sub

Private Sub tmdashonline_Click()
    DashBord_Online.Show
End Sub

Private Sub tmdataitem_Click()
DataItem_All.Show
End Sub

Private Sub tmdisccardprint_Click()
    PrintDiscBarcode.Show
End Sub

Private Sub tmdiscshop_Click()
    DiscoutFrm.Show 1
End Sub

Private Sub tmdocinout_Click()
    rbdoc_trans_brach2.Show 1
End Sub
Private Sub tmdocsend_Click()
    DocNotAccept.Show
End Sub
Private Sub tmdocsendFR_Click()
    DocNotAccept_FR.Show
End Sub
Private Sub tmfact_Click()
    FactFrm.Show 1
End Sub
Private Sub TMGRCUST_Click()
    'VsTBalCustGR.Show
End Sub
Private Sub tmfixcostsales_Click()
    
    AddLod_Data cusername, 0, " Ÿ»ÿ  ﬂ·›… «·„ÊœÌ·«  ", con
    Dim DDate1 As String, dDateSales As Date
    DDate1 = InputBox(" „‰  «—ÌŒ  ", , Format("01-01-" & Year(Date), "DD-MM-YY"))
    If DDate1 = "" Then Exit Sub
    If IsDate(DDate1) Then
        dDateSales = DateValue(DDate1)
    Else
        dDateSales = DateValue("01-01-" & Year(Date))
    End If
'    On Error GoTo myerror
    If MsgBox("Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  Ê «· ÕÊÌ·«  „‰ " & Format(dDateSales, "DD-MM-YYYY"), vbYesNo + vbDefaultButton2) = vbYes Then
        con.CommandTimeout = 5000
        cStr1 = " UPDATE    FILE6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE6_20.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE <= FILE6_20H.DATE ORDER BY DATE DESC ) " & _
                " FROM      FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO where (file6_20h.date) >= " & DateSq(dDateSales)
        con.Execute cStr1, nRec
        Inform "  „ Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  " & nRec
        If cBranch = "00" Then
            cStr1 = " UPDATE    FR6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST_FR WHERE FR6_20.ITEM = Q_ITEMCOST_FR.ITEM  AND FR6_20H.STORE = Q_ITEMCOST_FR.STORE AND Q_ITEMCOST_FR.DATE <= FR6_20H.DATE ORDER BY DATE DESC ) " & _
                    " FROM      FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO where (FR6_20H.date) >= " & DateSq(dDateSales)
            con.Execute cStr1, nRec
            Inform "  „ Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  «· ÊﬂÌ·« " & nRec
        End If
        If cBranch = "00" Then
            cStr1 = " UPDATE    FILE1_60 SET  cost = Coalesce((SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE1_60.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE <= FILE1_60H.DATE ORDER BY DATE DESC ),0) " & _
                    " FROM      FILE1_60 INNER JOIN FILE1_60H ON FILE1_60.DOC_NO = FILE1_60H.DOC_NO where (file1_60h.date) >= " & DateSq(dDateSales)
            con.Execute cStr1, nRec
            Inform "  „ Ÿ»ÿ  ﬂ·›… «· ÕÊÌ·«  1" & nRec
        
            cStr1 = " UPDATE    FILE1_60 SET  cost = Coalesce((SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE1_60.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE > FILE1_60H.DATE ORDER BY DATE ),0) " & _
                    " FROM      FILE1_60 INNER JOIN FILE1_60H ON FILE1_60.DOC_NO = FILE1_60H.DOC_NO where file1_60.cost = 0 and (file1_60h.date) >= " & DateSq(dDateSales)
            con.Execute cStr1, nRec
            Inform "  „ Ÿ»ÿ  ﬂ·›… «· ÕÊÌ·«  2" & nRec
        
            cStr1 = " UPDATE    FR1_60 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST_FR WHERE FR1_60.ITEM = Q_ITEMCOST_FR.ITEM AND Q_ITEMCOST_FR.DATE <= FR1_60H.DATE ORDER BY DATE DESC ) " & _
                    " FROM      FR1_60 INNER JOIN FR1_60H ON FR1_60.DOC_NO = FR1_60H.DOC_NO where (FR1_60h.date) >= " & DateSq(dDateSales)
            con.Execute cStr1, nRec
            Inform "  „ Ÿ»ÿ  ﬂ·›… «· ÕÊÌ·« " & nRec
        
        End If
        cStr1 = " UPDATE   FILE1_10 SET  COSTITEM = (SELECT TOP 1 ( COST ) FROM Q_ITEMCOST WHERE FILE1_10.ITEM = Q_ITEMCOST.ITEM ORDER BY DATE DESC ) From FILE1_10 "
        con.Execute cStr1, nRec
        Inform "  „ Ÿ»ÿ  ﬂ·›… 1 «·„ÊœÌ·«  " & nRec
        
        If cBranch = "00" Then
            cStr1 = " UPDATE   FILE1_10 SET  COSTITEM_FR = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST_FR WHERE FILE1_10.ITEM = Q_ITEMCOST_FR.ITEM ORDER BY DATE DESC ) From FILE1_10 "
            con.Execute cStr1, nRec
            Inform "  „ Ÿ»ÿ  ﬂ·›… 2 «·„ÊœÌ·«  " & nRec
        End If
        MsgBox "  „ Ÿ»ÿ «· ﬂ·›…"
        Exit Sub
    End If
myerror:
    MsgBox Err.Description
    Err.Clear

End Sub

Private Sub tmfrall_Click()
    VSTFRTOTAL.Show
End Sub

Private Sub tmfrgrtotal_Click()
    VsTFRGroup.Show
End Sub

Private Sub tmfrproft_Click()
    VSTFRPROFT.Show
End Sub

Private Sub tmfrsales_Click()
    VsTFRSales.Show
End Sub

Private Sub tmfrsalesday_Click()
    VsTFRBranshSalDay.Show
End Sub

Private Sub tmfrsuplsales_Click()
    VsTFRSalesSUPL.Show
End Sub

Private Sub tmgrbranch_Click()
    VsTBranshSalGroup.Show
End Sub

Private Sub tmgrmain_Click()
''ReDim aLocal(6)
''aLocal(0) = "FILE1_50G"
''aLocal(1) = "Code"
''aLocal(2) = "Desca"
''aLocal(3) = "ﬂÊœ "
''aLocal(4) = "»Ì«‰ "
''aLocal(5) = " ”ÃÌ· „Ã„Ê⁄«  —∆Ì”Ì…"
''aLocal(6) = 3
''FlagFrm2.bEdit = True
''FlagFrm2.myPublic = aLocal
''FlagFrm2.Show 1
MainGroup.Show 1
End Sub

Private Sub tmgroupsection_Click()
    VsTGroup.Show
End Sub
Private Sub tmimpcost_Click()
    impcostfrm.bEdit = True
    impcostfrm.Show
End Sub

Private Sub tmimportdata_Click()
    CopyData.Show
End Sub

Private Sub tminAGE_Click()
    VsAllTransage.Show
End Sub
Private Sub tmincome_Click()
    chargefrm.bEdit = True
    chargefrm.myPublic = 2
    chargefrm.Show
End Sub
Private Sub tmincomecode_Click()
chargeCodesfrm.myPublic = 2
chargeCodesfrm.bEdit = True
chargeCodesfrm.Show 1
End Sub
Private Sub tmincomemaincode_Click()
ReDim aPublic(6)
aPublic(0) = "FILE8_62"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ «·«Ì—«œ"
aPublic(4) = "»Ì«‰ «·«Ì—«œ"
aPublic(5) = "«Ì—«œ«  —∆Ì”Ì…"
aPublic(6) = 3
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aPublic
FlagFrm2.Show 1
End Sub
Private Sub tmindust_Click()
industFrm.Show 1
End Sub
Private Sub tminput_Click()
DamageFrm.myPublic = 2
DamageFrm.Show
End Sub

Private Sub tmitembranch_Click()
    VsTitemShop_br.Show
End Sub

Private Sub tmitembranch2_Click()
    VsTitemShop_itembr.Show
End Sub

Private Sub tminfact_Click()
    VsAllTransFact.Show
End Sub

Private Sub tminGROUP_Click()
    VsAllTransGROUP.Show
End Sub
Private Sub tmInput_ALL_Tab_Click()
    VsTInpit_Tab.Show
End Sub
Private Sub tminsec_Click()
    VsAllTranssec.Show
End Sub
Private Sub tmitemgroup_Click()
    GrItem.Show
End Sub
Private Sub tmitemgroupmain_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show
End Sub

Private Sub tmItemMainGroupMain_Click()

End Sub

Private Sub tmItemMainGroupRaw_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show
End Sub

Private Sub tmoutput_Click()
DamageFrm.myPublic = 3
DamageFrm.Show
End Sub

Private Sub tmitemshop_Click()
    VsTitemShop.Show
End Sub
Private Sub tmitemh_Click()
    VsTitem_H.Show
End Sub
Private Sub tmitemshopITEM_Click()
    VsTitemShop_item.Show
End Sub
Private Sub tmlookitem_Click()
    If cBranch = "00" Then
        ItemLookPh_m.Show
    Else
        ItemLookPh2.Show
    End If
End Sub

Private Sub tmmansales_Click()
    VsTSalesman.Show
End Sub
Private Sub tmmodelbal_Click()
    VsModelLook.Show
End Sub

Private Sub tmmansales2_Click()
    TDaySal_2.Show
End Sub

Private Sub tmmodelfact_Click()
    ShowModelFact.Show
End Sub

Private Sub tmmodelokaz_Click()
    VsOkazModel.Show
End Sub

Private Sub tmmodelokaz2_Click()
    VsOkaz_all.Show
End Sub

Private Sub tmmodelokazgr_Click()
    ModelOkazin.Show
End Sub
Private Sub tmmodels_Click()
    items.Show
End Sub
Private Sub tmmodelshow_Click()
    VsModelShow.Show
End Sub
Private Sub tmmodeltrans_Click()
    rbdoc_trans_brach.Show
End Sub
Private Sub tmmonthrep_Click()
    VsTSalesmonth2.Show
End Sub
Private Sub tmmosm_Click()
    MosmRate.Show 1
End Sub
Private Sub tmmosmgrbal_Click()
    VsTSEGR.Show
End Sub
Private Sub tmmosmsecbal_Click()
    VsTSECmosm.Show
End Sub

Private Sub tmmoveacc_Click()
    AccMove.Show
End Sub
Private Sub tmmovecust_Click()
    ClientMoveFrm.Show
End Sub

Private Sub tmmovesupp_Click()
    Supler_Move.Show
End Sub

Private Sub tmMoveTotAcc_Click()
    AccMove_Tot.Show
End Sub

Private Sub tmnewdoc_Click()
    NewDocFromFact.Show
End Sub

Private Sub tmNewDocBranch_Click()
    NewDocToBranch.Show
End Sub

Private Sub tmnewmosm_Click()
    FrmOpenMosm.Show 1

End Sub

Private Sub tmnewprice_Click()
    NewPrice.Show
End Sub

Private Sub tmonline_order_Click()
    OrderOnline_New.Show
End Sub

Private Sub tmopensal_Click()
    VsShip_Pay.Show
End Sub

Private Sub tmorder_rec_amazon_Click()
    VsOrder_Rec.Show
End Sub

Private Sub tmorderamazon_Click()
    OrderSouq.Show
End Sub
Private Sub tmOrderDem_ALL_Click()
    DemOrder_PurchItem.Show
End Sub
Private Sub tmOrderDem_Click()
    BalDemOrder_PurchItem.Show
End Sub

Private Sub tmOrderDOC_ALL_Click()
    DemOrder_PurchDoc.Show
End Sub

Private Sub tmorderpurch_Click()
    Ordere_Purch.Show
End Sub

Private Sub tmorderreponline_Click()
    RepOrderOnLIne.Show
End Sub

Private Sub tmpart_Click()
    Partfrm.bEdit = True
    Partfrm.Show
End Sub
Private Sub tmPath_Click()
    SettingFrm.Show 1
End Sub
Private Sub tmproduct_Click()
    productfrm.bEdit = True
    productfrm.Show
End Sub
Private Sub tmRawitem_Click()
    itemrawfrm.bEdit = True
    itemrawfrm.Show 1
End Sub
Private Sub tmsection_Click()
    sectionfrm.bEdit = True
    sectionfrm.Show
End Sub

Private Sub tmpart_code_Click()
ReDim aLocal(6)
aLocal(0) = "FILE8_71"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ «·‘—Ìﬂ"
aLocal(4) = "»Ì«‰ «·‘—Ìﬂ"
aLocal(5) = " ”ÃÌ· «·‘—ﬂ«¡"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show
End Sub

Private Sub tmpayshiponline_Click()
    Ship_Bills_PAY.Show
End Sub

Private Sub tmpaysupp_Click()
    VsPaySupl.Show
End Sub

Private Sub tmpirchokbr_Click()
    DocNotAccept_FR.Show

End Sub

Private Sub tmpriceokaz_Click()
    FixPriceOkaz.Show 1
End Sub

Private Sub tmprof1_Click()
frmProft.Show 1
End Sub

Private Sub tmproftbramonth_Click()
VsProftMonthBr.Show
End Sub

Private Sub tmproftrep_Click()
    frmProft.Show
End Sub

Private Sub tmproftsales_Click()
    grditemfrm2.Show
End Sub

Private Sub tmproftsupler_Click()
    VsTProftSuple.Show
End Sub

Private Sub tmPuechSUPear_Click()
    VsTPurchSup_YEAR.Show
End Sub
Private Sub tmpurchases_Click()
    purchasefrm.myPublic = 0
    purchasefrm.Show
End Sub
Private Sub tmrepdaysalestax_Click()
VsTDaytax.Show
End Sub
Private Sub tmrepsalestax_Click()
VsTSales_tax.Show
End Sub

Private Sub tmratemosm_Click()
    MosmRate.Show 1
End Sub

Private Sub tmrepsales_Click()
    TDaySal_2.Show
End Sub

Private Sub tmpurchbr_Click()
    purchasefrm.myPublic = 2
    purchasefrm.Show
End Sub

Private Sub tmpurchbrret_Click()
    purchasefrm.myPublic = 3
    purchasefrm.Show
End Sub

Private Sub tmpurchdoc_Click()
    VsTPurchdoc_no.Show
End Sub
Private Sub tmpurchmodel_Click()
    VsTPurchModel.Show
End Sub
Private Sub tmPURCHmonth_MODEL_Click()
    VsTSaleMODEL_YEAR.Show
End Sub

Private Sub tmPURCHmonth_year_Click()
    VsTPurchMonth_YEAR.Show
End Sub

Private Sub tmpurchordertot_Click()
'    BalDemOrder_Purch.Show
    BalDemOrder_Purch.Show

End Sub

Private Sub tmpurchSEC_Click()
    VsTPurchSec.Show
End Sub

Private Sub tmPurchSecYear_Click()
    VsTPurchSec_YEAR.Show
End Sub

Private Sub tmpurchstore_Click()
    VsTPurchStore.Show
End Sub

Private Sub tmPURCHstoreyear_Click()
    VsTPurchSTORE_YEAR.Show
End Sub

Private Sub tmpurchsupler_Click()
    VsTPurchSupler.Show
End Sub

Private Sub tmransbrancip_Click()
'    If UCase(RetSetting("COPY", App.Path & "\COPY.txt")) <> "COPY" Then
'        If cBranch < "60" Then
'            transdata_ip_br.Show 1
'        Else
        transdata_ip_br2.Show 1
'        End If
'    End If
End Sub
Private Sub tmrebranch_Click()
    sCatalog = RetSetting("Catalog", App.Path & "\conf.txt")
    sMdfName = sCatalog
    strCon = LoadConString
    strConfact = LoadConStringfact
    strConfact2 = LoadConStringfact2
    GetCon.Close
    OpenCon GetCon, strCon
    PassWord.Show
End Sub

Private Sub tmRecevdOrder_Click()
    sales_Recerv.Show
End Sub

Private Sub tmrecorder_supler_Click()
    Purch_Recerv.Show
End Sub

Private Sub tmrecordercust_Click()
    VsRepOrder.Show
End Sub

Private Sub tmrepalltrans_BR_Click()
    VsAllTransBRANCH.Show
End Sub

Private Sub tmrepalltrans_Click()
    VsAllTrans.Show
End Sub

Private Sub tmrepalltrans2_Click()
    VsAllTrans_2.Show
End Sub

Private Sub tmrepalltransGRMODEL_Click()
    VsAllTrans_GR.Show
End Sub

Private Sub tmrepbalfr_Click()
    VsTfrBranshSalModel.Show
End Sub

Private Sub tmrepbr1_Click()
    VsTBranshSalModel.Show
End Sub

Private Sub tmrepbr2_Click()
    VsTBranshSalSupl.Show
End Sub

Private Sub tmrepbr3_Click()
    VsTBranshSalDay.Show
End Sub

Private Sub tmrepbr5_Click()
    VsTBranshBALSupl.Show
End Sub

Private Sub tmrepchagebr_Click()
    rpCharge_br.Show
End Sub

Private Sub tmrepdemorder_Click()
    RepDemDoc.Show
End Sub
Private Sub tmrepeicefact_Click()
    RePriceFact.Show
End Sub

Private Sub tmrepinvoice_Click()
    RepTaxInvoice.Show
End Sub

Private Sub tmRepMonthTot_Click()
    Target_Month_All.Show
End Sub

Private Sub tmrepriceitem_Click()
    okaz_JUNIOR.Show
End Sub

Private Sub tmRepSales12_Click()
    VsTSales_Disc.Show
End Sub
Private Sub tmreptarget_Click()
    Target_Month_Branch.Show
End Sub
Private Sub tmrepvisa_Click()
    Rep_Visa.Show
End Sub
Private Sub tmret_sales_Click()
    Rep_Ret_Sales.Show
End Sub
Private Sub tmRETPURCH_Click()
    purchasefrm.myPublic = 1
    purchasefrm.Show
End Sub
Private Sub tmsalcustsub_Click()
    VsTBalCustSub.Show
End Sub

Private Sub tmsakesstore_Click()
    VsTSalesbarnch.Show
End Sub

Private Sub tmsalavgman_Click()
    VsTSalesmanavg.Show
End Sub
Private Sub tmsalday_Click()
    TDaySal.Show
End Sub

Private Sub tmsaldayfr_Click()
    TDaySal_FR.Show

End Sub

Private Sub tmsales_barcode_Click()
    VsTSales_BARCODE.Show
End Sub

Private Sub tmsales_Click()
'    salesfrm.lNotBranch = False
'    salesfrm.Show
sales_wholefrm.Show
End Sub
Private Sub tmsalesstore_Click()
VsTSalesStore.Show
End Sub

Private Sub tmsalesavgbranch_Click()
    VsTSalebranchavg.Show
End Sub

Private Sub tmsalesavgman_Click()
    VsTSalesmanavg.Show
End Sub

Private Sub tmsalesavgsec_Click()
    VsTSalesecahavg.Show
End Sub

Private Sub tmsalesavgsupler_Click()
    VsTSalesuplhavg.Show
End Sub

Private Sub tmSalesBranch_Click()
    salesfrm.lNotBranch = True
    salesfrm.Show
End Sub

Private Sub tmsalesday_Click()
    VsTSalesday.Show
End Sub

Private Sub tmsalesdiscdoc_Click()
    VsTSales_DISC_DOC.Show
End Sub
Private Sub tmsalesinvret_Click()
    salestAX.lIsRet = True
    salestAX.Show
End Sub

Private Sub tmsalespricedisctype_Click()
    VS_SALESPRICETYPE.Show
End Sub

Private Sub tmsalesrepall_Click()
    VSTSALES_ALL.Show
End Sub

Private Sub tmsalesretbr_Click()
    VSTSALESRETBR.Show
End Sub

Private Sub tmsalestax_Click()
    salestAX.lIsRet = False
    salestAX.Show
End Sub
Private Sub tmsalesfrmmonth_Click()
    VsTFRBranshSalMONTH.Show
End Sub

Private Sub tmsalesgr_Click()
    VsTSalesGtoup.Show
End Sub

Private Sub tmsalesmodel_Click()
    VsTSales_price.Show
End Sub

Private Sub tmsalesmonth_MODEL_Click()
    VsTSaleMODEL_YEAR.Show
End Sub

Private Sub tmsalesmonth_year_Click()
    VsTSalesMonth_YEAR.Show
End Sub

Private Sub tmsalesmonthbr_Click()
    VsTBranshSalMONTH.Show
End Sub

Private Sub tmsalesokaz_br_Click()
    VsTModelOkaz_br.Show
End Sub
Private Sub tmsalesokaz_Click()
    VsTModelOkaz.Show
End Sub

Private Sub tmsalesrepscal_Click()
    VsTSales_Scal.Show
End Sub

Private Sub tmsalesrepscalGR_Click()
    VsTSalesGr_Scal.Show
End Sub

Private Sub tmsalessecyear_Click()
    VsTSalessec_YEAR.Show
End Sub

Private Sub tmsalessSUPear_Click()
    VsTSalessup_YEAR.Show
End Sub

Private Sub tmsalesstoreyear_Click()
    VsTSalesSTORE_YEAR.Show
End Sub
Private Sub tmSALOKAZB_Click()
    VsTModelOkaz_B.Show
End Sub
Private Sub tmsec_Click()
'
'ReDim aLocal(6)
'aLocal(0) = "FILE1_10SC"
'aLocal(1) = "Code"
'aLocal(2) = "Desca"
'aLocal(3) = "ﬂÊœ ﬁ”„"
'aLocal(4) = "»Ì«‰ «·√ﬁ”«„"
'aLocal(5) = " ”ÃÌ· «·√ﬁ”«„"
'aLocal(6) = 2
'FlagFrm2.bEdit = True
'FlagFrm2.myPublic = aLocal
'FlagFrm2.Show 1

sectionCode.Show 1


End Sub

Private Sub tmsecbransh_Click()
    VsTBranshSalsec.Show
End Sub
Private Sub tmsecurity_Click()
    Security.Show 1
End Sub
Private Sub tmshopproft_Click()
Tproft_shop.Show
End Sub
Private Sub tmsenddataip_Click()
    transdata_ip.Show 1
End Sub
Private Sub tmsenddataip_fr_Click()
    transdata_ip_FR.Show 1
End Sub
Private Sub tmsexcode_Click()
    ReDim aLocal(6)
    aLocal(0) = "FILE0_81"
    aLocal(1) = "Code"
    aLocal(2) = "Desca"
    aLocal(3) = "ﬂÊœ «·‰Ê⁄"
    aLocal(4) = "»Ì«‰ "
    aLocal(5) = " ”ÃÌ· «·‰Ê⁄"
    aLocal(6) = 2
    FlagFrm2.bEdit = True
    FlagFrm2.myPublic = aLocal
    FlagFrm2.Show 1
End Sub

Private Sub tmshipcode_Click()
ReDim aLocal(6)
aLocal(0) = "SHIP"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ "
aLocal(4) = "»Ì«‰ "
aLocal(5) = "‘—ﬂ«  «·‘Õ‰"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1

End Sub

Private Sub tmStock_Click()
    StockFrm.Show
End Sub
Private Sub tt_item_Click()
    VsTItem.Show 1
End Sub
Private Sub tt_vend_Click()
VsSupp.Show 1
End Sub
Private Sub x_WAG_Click()
Load Emp
Emp.Show 1
End Sub
Private Sub XAbount_Click()
    Load About
    About.Show 1
End Sub
Private Sub xCASH1_Click()
    publicFlag = 4
    Load Vs_Cash
    Vs_Cash.Show 1
End Sub

Private Sub tmstockitemdoc_Click()
    StockAllDoc.Show
End Sub
Private Sub tmstores_Click()
    StoreData.Show
End Sub
Private Sub tmsubuser_Click()
    SubUser.Show
End Sub
Private Sub tmsuplerfrsales_Click()
    VsTFRBranshSalSupl.Show
End Sub
Private Sub tmsuppgr_Click()
ReDim aLocal(6)
aLocal(0) = "FILE4_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ „Ã„Ê⁄…"
aLocal(4) = "»Ì«‰ „Ã„Ê⁄« "
aLocal(5) = "„Ã„Ê⁄«  „Ê—œÌ‰"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub tmsuppl_Click()
    suplerfrm.Show
End Sub
Private Sub tmsuppmove_Click()
    supMovefrm.Show
End Sub
Private Sub tmTCUST_Click()
    VsTBalCust.Show
End Sub
Private Sub tmtbalscal_Click()
    VsModelScal.Show
End Sub

Private Sub tmTargetBr_Click()
    Target_Branch.Show
End Sub

Private Sub tmtrans_br_Click()
    trans_FR.bEdit = True
    trans_FR.Show
End Sub

Private Sub tmtrans_Click()
    transfrm.bEdit = True
    transfrm.Show
End Sub

Private Sub tmtransout_Click()
'    Transfrm.myPublic = 1
    transfrm.Show
End Sub

Private Sub tmtransacc_Click()
    AccDoc.Show
End Sub
Private Sub tmtransbarcode_Click()
    ConvertBarcode.Show 1
End Sub

Private Sub tmtranscode2_Click()
    boxtrans_2.bEdit = True
    boxtrans_2.Show
End Sub

Private Sub tmtranscost_Click()
    VsTcosttrans1.Show
End Sub

Private Sub tmtransdata_Click()
'    If cBranch = "00" Then
'        transdata_mainfrm.Show 1
'    Else
'        If InputBox("", "") = "654321" Then
'            transDatafrm.Show 1
'        End If
'    End If
End Sub

Private Sub tmTRANSFROMIP_FR_Click()
    If InputBox("", "") = "258" Then transdataFR_ip.Show
End Sub

Private Sub tmtranstocust_Click()
    If cBranch = "00" Then
        TransCust_Main.Show 1
    Else
        TransFromCust.Show 1
    End If
End Sub

Private Sub tmtransmosm_Click()
    TransMosm.myPublic = 2
    TransMosm.Show
End Sub

Private Sub tmtsales_Click()
    VsTSales.Show
End Sub

Private Sub tmvsstore_Click()
    VsTStore.Show
End Sub

Private Sub tsalescust2_Click()
VsTBalCustSub.Show
End Sub

Private Sub tmtsalesFACT_Click()
    VsTSalesFact.Show
End Sub

Private Sub tmtsalesAGE_Click()
    VsTSalesAGE.Show
End Sub

Private Sub tmtsalessec_Click()
    VsTSalessec.Show
End Sub

Private Sub tmtsalesSUPL_Click()
    VsTSalesSUPL.Show
End Sub

Private Sub tmvisa_Click()
    VisaFrm.Show
End Sub

Private Sub tmzone_Click()
ReDim aLocal(6)
aLocal(0) = "BRANCH_GR"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ "
aLocal(4) = "»Ì«‰ «·„‰«ÿﬁ"
aLocal(5) = " ”ÃÌ· «·„‰«ÿﬁ"
aLocal(6) = 1
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub

Private Sub trmtrans2_Click()
    trans_2.nFlagPublic = 0
    trans_2.Show
End Sub
Private Sub trmtrans3_Click()
    trans_2.nFlagPublic = 1
    trans_2.Show
End Sub

Private Sub xCASH2_Click()
'Cashfrm.bEdit = True
    Cashfrm.myPublic = 2
    Cashfrm.Show
End Sub
Private Sub xCen_Cost_Click()
publicFlag = 5
Load Flag
Flag.Show 1
End Sub
Private Sub xprofff_Click()
    Load Profit
    Profit.Show 1
End Sub

Private Sub xchksupp_Click()
    chq2.Show 1
End Sub
Private Sub xfolcost_Click()
    VsPrice2.Show 1
End Sub
Private Sub XFOLPRICE_Click()
    VsPrice.Show 1
End Sub

Private Sub xGrVEND_Click()
supMovefrm.Show
End Sub

Private Sub xmacc_Click()
    frmProft.Show 1
End Sub
Private Sub XMBALBOX_Click()
    BalBox.Show 1
End Sub
Private Sub XMBALSTORE_Click()
    StoreBal.Show 1
End Sub
Private Sub XMBARPRINT_Click()
     Morsh_Bar.Show
End Sub

Private Sub xItemMoveFR_Click()
    ModelMove_FR.Show
End Sub

Private Sub xmcash3_Click()
    bEdit = True
    Cashfrm.myPublic = 4
    Cashfrm.Show
End Sub
Private Sub xmcash4_Click()
'Cashfrm.bEdit = True
    Cashfrm.myPublic = 3
    Cashfrm.Show
End Sub
Private Sub XMCODEBOX_Click()
    MonyBox.Show 1
End Sub
Private Sub XMPAY_Click()
    Vs_Pay.Show 1
End Sub

Private Sub XMREDEM_Click()
    VsSal.Show 1
End Sub
Private Sub xMReSTORE_Click()
VsSal.Show 1
End Sub
Private Sub XMREDEM2_Click()
    Redem.Show 1
End Sub

Private Sub xMSal_Click()
    ManFrm.Show 1
End Sub
Private Sub XMSALCODE_Click()
Vstsalsupp.Show 1
End Sub
Private Sub XMSTORE_Click()
ReDim aLocal(6)
aLocal(0) = "FILE0_40"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ «·„Œ“‰"
aLocal(4) = "»Ì«‰ «·„Œ“‰"
aLocal(5) = " ”ÃÌ· « ·„Œ«“‰"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub XMTCUSTSAL_Click()
    F_SalCust.Show 1
End Sub
Private Sub XMTRANS_Click()
    transfrm.Show
End Sub
Private Sub XMTSTORE_Click()
    VsStore.Show 1
End Sub
Private Sub XMVISA_Click()
    SalVisa.Show 1
End Sub
Private Sub XRETPURCH_Click()
    purchasefrm.myPublic = 1
    purchasefrm.Show 1
End Sub
Private Sub storerep_Click()
    rpItem.Show
End Sub
Private Sub XBDATA_Click()

itemsfrm.Show
End Sub
Private Sub xCASH3_Click()
    publicFlag = 3
    Load Vs_Cash
    Vs_Cash.Show 1
End Sub
Private Sub xCASH4_Click()
    publicFlag = 2
    Load Vs_Cash
    Vs_Cash.Show 1
End Sub
Private Sub xCashed_Click()
'Cashfrm1.bEdit = True
'Cashfrm1.Show
'Cashfrm.bEdit = True
    Cashfrm.myPublic = 1
    Cashfrm.Show
End Sub
Private Sub XcHARG_Click()
    publicFlag = 1
    Load Vs_Charg
    Vs_Charg.Show 1
End Sub
Private Sub XCHARGE_Click()
    publicFlag = 7
    Load Flag
    Flag.Show 1
End Sub
Private Sub XCHQ_Click()
    publicFlag = 1
    Chq.Show 1
End Sub
Private Sub xClientData_Click()
Clients.myFlag = 1
Clients.Show 1
End Sub
Private Sub xClientGroup_Click()
ReDim aLocal(6)
aLocal(0) = "FILE3_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ"
aLocal(4) = "≈”„ «·„Ã„Ê⁄…"
aLocal(5) = " ”ÃÌ· «·⁄„·«¡"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub xClientMove_Click()
    ClientMoveFrm.Show
End Sub
Private Sub xClientReport_Click()
Load ClientReports
ClientReports.Show 1
End Sub
Private Sub xCompGroup_Click()
End Sub
Private Sub xCredit_Click()
    Load CREDIT
    CREDIT.Show 1
End Sub
Private Sub XCREDITREP_Click()
    Load RepCredit
    RepCredit.Show 1
End Sub
Private Sub XDATA_Click()
Load BANK
BANK.Show 1
End Sub
Private Sub xDel_Move_Click()
If MsgBox("«·€«¡ﬂ· «·»Ì«‰«   : Â· «‰  „Ê«›ﬁ ø", 4 + 256) <> 6 Then Exit Sub
Dim aString(12)
aString(0) = "file0_10"
aString(1) = "file1_10"
aString(2) = "file1_11"
aString(3) = "file1_30"
aString(4) = "file1_60"
aString(5) = "file1_70"
aString(6) = "file1_50"
aString(7) = "file3_10"
aString(8) = "file4_10"
aString(9) = "file6_10"
aString(10) = "file6_20"
aString(11) = "file7_20"
For i = 0 To UBound(aString) - 1
    cString = "Delete Distinctrow " & aString(i) & ".*" & " From " & aString(i)
    mydb.Execute cString
Next
End Sub
Private Sub xExit_Click()
    End
End Sub
Private Sub xGroup_Click()
Load ItemsGrp
ItemsGrp.Show 1
End Sub
Private Sub xIMPORT_Click()
Load Import
Import.Show 1
End Sub
Private Sub xinventory_Click()
'Load inventory
'inventory.Show 1
Vs_Stock.Show 1
End Sub
Private Sub xItemMove_Click()
    ModelMove.Show
End Sub
Private Sub xMainGroup_Click()
publicFlag = 2
Load Flag
Flag.Show 1
End Sub
Private Sub xMortal_Click()
publicFlag = 2
Load Vs_Input
Vs_Input.Show 1
End Sub
Private Sub xOutPut_Click()
publicFlag = 1
Load Vs_Input
Vs_Input.Show 1
End Sub
Private Sub xpurchases_Click()
purchasefrm.myPublic = 0
purchasefrm.Show 1
End Sub
Private Sub xRepSAlMan_Click()
    Load rsalman
    rsalman.Show 1
End Sub
Private Sub xRetSales_Click()
salesfrm.myPublic = 1
salesfrm.Show 1
End Sub
Private Sub xSales_Click()
salesfrm.myPublic = 0
salesfrm.Show 1
End Sub
Private Sub xStoreMove_Click()
Load StoreMove
StoreMove.Show 1
End Sub
Private Sub xStores_Click()
publicFlag = 1
Load Flag
Flag.Show 1
End Sub
Private Sub xStoreTrans_Click()
    Load Vs_Trans
    Vs_Trans.Show 1
End Sub
Private Sub xutil11_Click()
Load Utility
Utility.Show 1
End Sub

Private Sub XTMBALSUPP_Click()
    VsTCustSales.Show
End Sub
Private Sub xVendorData_Click()
supfrm.Show 1
End Sub
Private Sub xVendorGroup_Click()
ReDim aLocal(6)
aLocal(0) = "FILE4_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ"
aLocal(4) = "≈”„ «·„Ã„Ê⁄…"
aLocal(5) = " ”ÃÌ· «·„Ê—œ"
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
End Sub
Private Sub FixData1()
Dim FS1 As New ADODB.command
OpenCon con
On Error Resume Next


cString = "ALTER TABLE [dbo].[FILE6_20] ADD [S_OKAZ] [decimal](18, 2) NULL "

FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
closeCon con
Err.Clear
End Sub
Private Sub FixData2()
Dim FS1 As New ADODB.command, cDesca As String
cDesca = "’·«ÕÌ… „— Ã⁄"
OpenCon con
On Error Resume Next
If GetDesca("SELECT ID FROM [option] WHERE ID = 7 ", con) = "" Then
    cStr1 = " INSERT INTO [option] ([desca]) VALUES ( " & addstring(cDesca) & " ) "
    con.Execute cStr1
End If
Err.Clear
End Sub

Private Sub FixData4()
Dim FS1 As New ADODB.command
OpenCon con
On Error Resume Next
cString = "alter TABLE [dbo].[FILE0_10H] add   [branch] [nvarchar](50) COLLATE Arabic_CI_AS NULL"

FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
closeCon con
Err.Clear
End Sub

Private Sub FixData12()
Dim FS1 As New ADODB.command
OpenCon con
On Error Resume Next
cString = "ALTER TABLE [dbo].[FILE7_20h] ADD [isnew2] [bit] NULL"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
closeCon con
Err.Clear
End Sub
Private Sub FixData13()
Dim FS1 As New ADODB.command
OpenCon con
On Error Resume Next
cString = "ALTER TABLE [dbo].[FILE7_20h] ADD [isnew3] [bit] NULL"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
closeCon con
Err.Clear
End Sub

Private Sub FixData5()
Dim FS1 As New ADODB.command
OpenCon con
On Error Resume Next

cString = " ALTER VIEW [dbo].[FILE1_10H] AS SELECT     model, MAX(desca) AS DESCA, MAX([GROUP]) AS [GROUP], MAX(RATE) AS RATE, MAX(OKAZ) AS OKAZ, MAX(FACT) AS FACT, MAX(MOSM) AS MOSM, MAX(modelno)  AS MODELNO, MAX(MODELFACT) AS MODELFACT, MAX(code) AS code, MAX(SUPP) AS SUPP, MAX(MODELFACT0) AS MODELFACT0, REDEM, FIXPRICE,  SHOWSALES, ISOKAZITEM, SECTION " & _
             " FROM         dbo.FILE1_10 GROUP BY model, REDEM, FIXPRICE, SHOWSALES, ISOKAZITEM, SECTION "

FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
closeCon con
Err.Clear
End Sub




Sub FixStoreCode()
'    con.Execute " UPDATE FILE0_40   SET CODE = '0' + FILE0_40.CODE WHERE LEN(FILE0_40.CODE) = 1 "
'    con.Execute " UPDATE FILE1_60H  SET STORE1 = '0' + FILE1_60H.STOTE1 WHERE LEN(STORE1) = 1 "
'    con.Execute " UPDATE FILE1_60H  SET STORE2 = '0' + FILE1_60H.STOTE2 WHERE LEN(STORE2) = 1 "
'    con.Execute " UPDATE FILE7_20H  SET STORE= '0' + FILE7_20H.STOREWHERE LEN(FILE0_40.CODE) = 1 "


End Sub
Sub Show_NewDoc()
    Set grid1.DataSource = data1
    data1.connectionString = strCon
    grid1.Rows = 1
    fixGrd
    myload
    
End Sub
Private Sub myload()
    Dim cWhere As String
    cString = " SELECT     flag, desca, DOC_NO, DATE, [name], QUANT, USERSEND, TIMESEND FROM         Q_DATA_SEND "
    data1.RecordSource = cString
    data1.Refresh
    fixGrd
End Sub
Sub fixGrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 10
    .RowHeight(0) = 1000
    .TextMatrix(0, 0) = ""
    .TextMatrix(0, 1) = "‰Ê⁄ «·„” ‰œ"
    .TextMatrix(0, 2) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 3) = " «—ÌŒ"
    .TextMatrix(0, 4) = "«·≈”„"
    .TextMatrix(0, 5) = "⁄œœ «·ﬁÿ⁄"
    .TextMatrix(0, 6) = "≈”„ «·—«”·"
    .TextMatrix(0, 7) = " «·Êﬁ "
    .TextMatrix(0, 8) = " › Õ «·„” ‰œ"
    .TextMatrix(0, 9) = "Exit"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 2000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1500
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    
    .ColWidth(6) = 2500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1000
    .ColWidth(9) = 800
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTDate
    .ColDataType(4) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColComboList(8) = "..."
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub checkCopy()
On Error GoTo myerror
Dim aDrive As Variant, fs As New FileSystemObject, cBackDrive As String
cBackDrive = GetDesca("SELECT BACK FROM ADDRESS", con)
aDrive = aLastDrive(False)
cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
cFileName = cDir & "\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak"
If Not fs.FileExists(cFileName) Then
    MsgBox "”Ì „ ⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
    If DoBackUp Then Inform " „ ⁄„· ‰”Œ… »‰Ã«Õ"
End If
Exit Sub
myerror:
'MsgBox Err.Description
Err.Clear
End Sub
Private Function DoBackUp(Optional pFile As String = "") As Boolean
'On Error GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, aDrive As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
aDrive = aLastDrive(False)
cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
Dim sMsg As String
MyCreateFolder (cDir)

cFileName = cDir & "\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak"
If createBackUp(cFileName) Then DoBackUp = True
'fs.CopyFile App.Path & "\databackup\back_SHOP.bak", cFileName
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function createBackUp(pFileName) As Boolean
Dim cFile As String
Dim con As New ADODB.Connection
OpenCon con

Dim cmd
'Set cmd = con.CreateObject("ADODB.Command")
'cmd.CommandTimeout = 120   ' number of seconds
Dim FS1 As New ADODB.command
FS1.CommandType = adCmdText


Set FS1.ActiveConnection = con
cString = "BACKUP DATABASE " & sCatalog & " TO  DISK = N'" & pFileName & "' WITH  RETAINDAYS = 1, NOFORMAT, INIT,  NAME = N'over-Full Database Backup', SKIP,  NOREWIND, NOUNLOAD, STATS = 10"
FS1.CommandText = cString
FS1.CommandTimeout = 6000
FS1.Execute
Set FS1 = Nothing
closeCon con
createBackUp = True
End Function
Sub ClosedCashDoc()
    OpenCon con
    Dim pDate As Date
'    con.Execute " UPDATE FILE7_10H SET CLOSED = 1   WHERE CLOSED  = 0 AND DATE < " & DateSq(Date)
'    con.Execute " UPDATE FILE7_20H SET CLOSED = 1   WHERE CLOSED  = 0 AND DATE < " & DateSq(Date)
'    con.Execute " UPDATE FILE1_60H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
    
    If lIsBranchStore Then
        con.Execute " UPDATE FILE8_50H SET ISCLOSED = 1 WHERE BRANCH = " & MyParn(cBranch) & " AND ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_60H SET ISCLOSED = 1 WHERE BRANCH = " & MyParn(cBranch) & " AND ISCLOSED  = 0 AND DATE < " & DateSq(Date)
'        con.Execute " UPDATE FILE8_70H SET ISCLOSED = 1 WHERE BRANCH = " & MyParn(cBranch) & " AND ISCLOSED  = 0 AND DATE < " & DateSq(Date)
'        con.Execute " UPDATE FILE0_51  SET ISCLOSED = 1 WHERE BRANCH = " & MyParn(cBranch) & " AND ISCLOSED  = 0 AND DATE < " & DateSq(Date)
'        con.Execute " UPDATE FILE0_52  SET ISCLOSED = 1 WHERE BRANCH = " & MyParn(cBranch) & " AND ISCLOSED  = 0 AND DATE < " & DateSq(Date)
    Else
        con.Execute " UPDATE FILE8_00H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_10H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_20H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_30H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_40H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_50H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_60H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE8_70H SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
        con.Execute " UPDATE FILE0_51  SET ISCLOSED = 1 WHERE ISCLOSED  = 0 AND DATE < " & DateSq(Date)
    End If

End Sub
Sub FixData_ALL_FR()
Dim cString As String
OpenCon con
On Error Resume Next
Inform cBranch

'cString = "ALTER TABLE [dbo].[branch] ADD  [VISA] [nvarchar](50) COLLATE Arabic_CI_AS NULL, [IPSERVER] [nvarchar](50) COLLATE Arabic_CI_AS NULL, [DATASERVER] [nvarchar](50) COLLATE Arabic_CI_AS NULL"
'createCommand cString, con
'
'
'cString = "ALTER TABLE [dbo].[FILE6_20H] ADD     [BOXVISA] [nvarchar](50) NULL,     [codeVISA] [int] NULL"
'createCommand cString, con
'
'cString = "CREATE TABLE [dbo].[VISA](  [CODE] [int] NULL,  [DESCA] [nvarchar](50) NULL,    [BANK] [nvarchar](6) NULL,  [rate] [decimal](18, 4) NULL,   [branch] [nvarchar](2) NULL) ON [PRIMARY]"
'createCommand cString, con
'
'cString = "ALTER TABLE [dbo].[FILE6_25] ADD [branch] [nvarchar](2) COLLATE Arabic_CI_AS NULL"
'createCommand cString, con
'
'cString = ReadFile(App.Path & "\TXT\QSALESDTL.TXT")
'createCommand cString, con
'
'cString = "ALTER TABLE [dbo].[menu] ALTER COLUMN [DescA] [nvarchar](100) COLLATE Arabic_CI_AS NULL"
'createCommand cString, con
'
'cString = "ALTER TABLE [dbo].[FILE4_10] ADD [ISOPEN] [bit] NOT NULL CONSTRAINT [DF_FILE4_10_ISOPEN]      DEFAULT ((0)) "
'createCommand cString, con
'
'cString = "ALTER TABLE [dbo].[SUBCUST] ADD [ISNEW] [bit] NOT NULL CONSTRAINT [DF_SUBCUST_ISNEW]      DEFAULT ((1)) "
'createCommand cString, con
'
'con.Execute " UPDATE FILE4_10 SET isopen = 0 "
'
'cString = ReadFile(App.Path & "\TXT\Qfile4_11disc.TXT")
'createCommand cString, con
'
''cString = ReadFile(App.Path & "\TXT\QFILE4_11.TXT")
''createCommand cString, con
'
'
'cString = " ALTER TABLE [dbo].[FILE8_50H] ALTER COLUMN [DOC_NO] [varchar](8) COLLATE Arabic_CI_AS NOT NULL"
'createCommand cString, con
'
'cString = " ALTER TABLE [dbo].[FILE8_60H] ALTER COLUMN [DOC_NO] [varchar](8) COLLATE Arabic_CI_AS NOT NULL"
'createCommand cString, con
'
'cString = " ALTER TABLE [dbo].[FILE8_60] ALTER COLUMN [DOC_NO] [varchar](8) COLLATE Arabic_CI_AS NOT NULL"
'createCommand cString, con
'
'cString = " alter table [dbo].[ADDRESS] add   [branch] [nvarchar](50) NULL"
'createCommand cString, con
'
'cString = " alter table [dbo].[FR7_20H] add   [Notes_Doc] [nvarchar](50) NULL"
'createCommand cString, con
'
'cString = " alter table [dbo].[FR7_10H] add   [Notes_Doc] [nvarchar](50) NULL"
'createCommand cString, con
'
'cString = " ALTER TABLE [dbo].[VISA] ADD [ISSTOP] [bit] NOT NULL"
'createCommand cString, con
'con.Execute " UPDATE VISA SET ISSTOP = 0 WHERE ISSTOP IS NULL "
'
'If cBranch > "60" Then
'    con.Execute " update ADDRESS set branch = " & addstring(cBranch)
'End If
closeCon con
Err.Clear
End Sub
Private Sub CopyServer()
'On Error GoTo myerror
Dim aDrive As Variant, fs As New FileSystemObject, cFileName As String
aDrive = aLastDrive(False)
cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
cFileName = App.Path & "\databackup\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak"
If Not fs.FileExists(cFileName) Then
    MsgBox "”Ì „ ⁄„· ‰”Œ… «Õ Ì«ÿÌ… " & cFileName
    If DoBackUp(cFileName) Then Inform " „ ⁄„· ‰”Œ… »‰Ã«Õ"
    If Not fs.FileExists(cFileName) Then
        MsgBox "·„ Ì „ ⁄„· ‰”Œ… „‰ «·»Ì«‰«  ·⁄–« «·ÌÊ„" & cFileName
    End If
End If
Exit Sub
myerror:
'MsgBox Err.Description
Err.Clear
End Sub
Function LoadConStringServer()
Dim cServerName As String, cUserId As String, cPassword As String, cCatalog2 As String
If lServerOnLine Then
'    cServerName = "junior-sql.database.windows.net"
'    cUserId = decrypt("062F8C407C77", "dr")
'    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
Else
    cServerName = RetSetting("server", App.Path & "\conf.txt")
    cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
    cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
End If


cCatalog2 = RetSetting("CATALOG", App.Path & "\conf.txt")
LoadConStringServer = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
            & "catalog=" & cCatalog2 & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Sub FixData_ALL()
Dim cString As String
OpenCon con
On Error Resume Next


'cString = " ALTER TABLE [dbo].[FILE6_20] ADD    [S_OKAZ_2] [decimal](18, 2) NULL "
'createCommand cString, con
'
'    cString = " DROP TABLE Q_RETBAL_ITEM"
'    createCommand cString, con
'''''
'    cString = " CREATE TABLE [dbo].[Q_RETBAL_ITEM]( [DOC_NO] [varchar](12) NOT NULL,    [store] [nvarchar](3) NULL, [date] [datetime] NULL, [ITEM] [int] NULL,  [QUANT] [decimal](18, 2) NULL,  [QUANT_RET] [int] NOT NULL, [PRICE] [numeric](38, 6) NULL,  [TOTAL] [numeric](38, 5) NULL,  [rate] [decimal](18, 2) NOT NULL,   [PRICE_SALES] [numeric](38, 6) NULL,    [MAN] [nvarchar](5) NULL,   [DOC_DISC] [nvarchar](6) NULL) ON [PRIMARY]"
'    createCommand cString, con
'''''
'    cString = " CREATE PROCEDURE [dbo].[retDOC_NO]  @DOC_NO  nVarChar(15), @RETURN NVARCHAR(15)  OUTPUT AS BEGIN DECLARE @MYDOC_NO NVARCHAR(15) SELECT @MYDOC_NO = QFILE6_20.DOC_NO  From QFILE6_20 where QFILE6_20.DOC_NO  = @DOC_NO  SET @RETURN = @MYDOC_NO END"
'    createCommand cString, con
'''''
'    cString = " CREATE PROCEDURE [dbo].[SALES_DOC_FIND]     @DOC_NO NVARCHAR(15) AS BEGIN   SELECT  QFILE6_20H.* From QFILE6_20H WHERE DOC_NO = @DOC_NO END"
'    createCommand cString, con
'''''
'    cString = " ALTER TABLE [dbo].[Q_RETBAL_ITEM] ADD   [S_OKAZ] [int] NULL"
'    createCommand cString, con
'''''
'    cString = "ALTER TABLE [dbo].[FILE1_10]       ADD  [ISONLINE] [bit] NOT NULL CONSTRAINT [DF_FILE1_10_ISONLINE] DEFAULT ((0)) "
'    createCommand cString, con
'''''
'    cString = "ALTER TABLE [dbo].[FILE1_10H]       ADD  [ISONLINE] [bit] NOT NULL CONSTRAINT [DF_FILE1_10H_ISONLINE] DEFAULT ((0)) "
'    createCommand cString, con
'''''
'    cString = "ALTER TABLE [dbo].[FILE0_40]       ADD  [online] [bit] NOT NULL CONSTRAINT [DF_FILE0_40_online] DEFAULT ((0)) "
'    createCommand cString, con
'
'    cString = " ALTER TABLE [dbo].[ADDRESS] ADD     [DATE1_DISC] [nvarchar](50) NULL,   [DATE2_DISC] [nvarchar](50) NULL,   [DISC] [numeric](18, 2) NULL     "
'    createCommand cString, con
'''''
'
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [note_ret] [nvarchar](100) NULL     "
'    createCommand cString, con
'''''
'
'
'    cString = "ALTER TABLE [dbo].[FILE0_50]  ADD  [ISBOXONLINE] [bit] NOT NULL CONSTRAINT [DF_FILE0_50_ISBOXONLINE] DEFAULT ((0)) "
'    createCommand cString, con
'
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [ONLINE_DOC] [nvarchar](10) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [Shipping_City] [nvarchar](1000) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CITY] [nvarchar](1000) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [STREET] [nvarchar](1000) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [Payment_Method] [nvarchar](100) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20] add   [DESCA2] [nvarchar](1000) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20] add   [SKU] [nvarchar](20) NULL     "
'    createCommand cString, con
'
'    If cBranch = "00" Then
'        cString = " alter TABLE [dbo].[FILE6_90H] add   [SHIP_NO] [nvarchar](100) NULL     "
'        createCommand cString, con
'    End If
'
'    cString = " CREATE  PROCEDURE [dbo].[Ret_Phone]     @PHONE Varchar(15),     @RETURN NVARCHAR(200)  OUTPUT AS BEGIN  DECLARE @DESCA NVARCHAR(200)    SELECT TOP 1 @DESCA = DESCA From SUBCUST where PHONE  = @PHONE  SET @RETURN = @DESCA END    "
'    createCommand cString, con
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [ISNODEL] [bit] NOT NULL CONSTRAINT [DF_FILE6_20H_ISNODEL]  DEFAULT ((0))     "
'    createCommand cString, con
'
'    cString = " ALTER TABLE [dbo].[FILE0_40] ADD   [ISNOGRDISC] [bit] NOT NULL CONSTRAINT [DF_FILE0_40_ISNOGRDISC]  DEFAULT ((0))     "
'    createCommand cString, con
'
'    cString = " ALTER TABLE [dbo].[file6_20h] ADD [NORET_CHARGE1] [bit] NOT NULL CONSTRAINT [DF_file6_20H_NORET_CHARGE1] DEFAULT ((0))"
'    createCommand cString, con
'
'    cString = ReadFile(App.Path & "\TXT\Q_SHIP.TXT")
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE7_20H] add   [Notes] [nvarchar](100) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE7_10H] add   [Notes] [nvarchar](100) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_DISC] [nvarchar](10) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_DOC] [nvarchar](6) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_DESCA] [nvarchar](100) NULL     "
'    createCommand cString, con
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_VALUE] [numeric](18, 2) NULL   "
'    createCommand cString, con
'
'    If cBranch = "00" Then
'        cString = " alter TABLE [dbo].[FR6_20H] add   [CARD_DISC] [nvarchar](10) NULL     "
'        createCommand cString, con
'
'        cString = " alter TABLE [dbo].[FR6_20H] add   [CARD_DOC] [nvarchar](6) NULL     "
'        createCommand cString, con
'
'        cString = " alter TABLE [dbo].[FR6_20H] add   [CARD_DESCA] [nvarchar](100) NULL     "
'        createCommand cString, con
'
'        cString = " alter TABLE [dbo].[FR6_20H] add   [CARD_VALUE] [numeric](18, 2) NULL   "
'        createCommand cString, con
'    End If
'
'    cString = " ALTER TABLE [dbo].[FILE6_20] ADD    [TIME_ADD] [datetime] NOT NULL CONSTRAINT [DF_FILE6_20_TIME_ADD]  DEFAULT (getdate())     "
'    createCommand cString, con
'
'    cString = "ALTER TABLE [dbo].[SUB_USER]       ADD  [ISRETURN] [bit] NOT NULL CONSTRAINT [DF_SUB_USER_ISRETURN] DEFAULT ((0)) "
'    createCommand cString, con
'
'    If lServerOnLine And cBranch = "00" Then
'        cString = " CREATE view  [dbo].[QFILE7_20H] AS " & _
'                    " SELECT        FILE7_20H.DOC_NO, FILE7_20H.store, FILE7_20H.date, SUM(FILE7_20.QUANT) AS QUANT , SUM(FILE7_20.TOTAL) AS TOTAL , FILE7_20H.discount, FILE7_20H.code, 1 AS iSBRANCH  , inv_no , FILE7_20H.MOSM AS MOSMINV , TAX FROM            FILE7_20H INNER JOIN FILE7_20 ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO GROUP BY FILE7_20H.DOC_NO, FILE7_20H.store, FILE7_20H.date, FILE7_20H.discount, FILE7_20H.code, file7_20H.DISCOUNT , inv_no , FILE7_20H.MOSM , TAX " & _
'                    " UNION ALL SELECT        FILE7_10H.DOC_NO, FILE7_10H.store, FILE7_10H.date,  SUM(FILE7_10.QUANT * -1 ) AS QUANT , SUM(FILE7_10.TOTAL * -1 ) AS TOTAL , FILE7_10H.discount * -1 , FILE7_10H.code, 1 AS iSBRANCH  , inv_no , FILE7_10H.MOSM AS MOSMINV , TAX * -1  AS TAX FROM            FILE7_10H INNER JOIN FILE7_10 ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO GROUP BY FILE7_10H.DOC_NO, FILE7_10H.store, FILE7_10H.date, FILE7_10.ITEM, FILE7_10H.discount, FILE7_10H.code, file7_10H.DISCOUNT , inv_no , FILE7_10H.MOSM , TAX"
'    createCommand cString, con
'    End If
'
'    cString = " alter TABLE [dbo].[FR6_20] add   [MAN2] [nvarchar](10) NULL     "
'    createCommand cString, con
'    cString = " alter TABLE [dbo].[FILE6_20] add   [MAN2] [nvarchar](10) NULL     "
'    createCommand cString, con
'
'    If cBranch <> "00" And Not lServerOnLine Then
'        cString = " alter TABLE [dbo].[file1_10] add   [Outlet] [bit] NOT NULL CONSTRAINT [DF_1_10_Outlet] DEFAULT ((0))  "
'        createCommand cString, con
'
'        cString = "  ALTER VIEW [dbo].[QFILE6_20] AS  SELECT        FILE6_20H.DOC_NO, FILE6_20H.store, FILE6_20H.date, FILE6_20H.CASH, FILE6_20H.VISA, FILE6_20H.BOX, FILE6_20H.PAY, FILE6_20H.branch, FILE6_20H.username_disc, FILE6_20.ITEM, FILE6_20.QUANT, (FILE6_20.PRICE ) AS PRICE, (FILE6_20.TOTAL ) AS TOTAL, FILE6_20H.discount, FILE6_20H.rate, FILE6_20.cost, FILE6_20.PRICE_C, FILE6_20.MAN, FILE6_20.ISDISC2_1,                          FILE6_20.S_OKAZ, FILE1_10.MODEL, FILE1_10.[GROUP], FILE1_10.code, FILE1_10.MOSM, FILE1_10.SECTION, 1 AS iSBRANCH, FILE6_20H.SALES_RET, FILE6_20H.username_ret, FILE6_20H.PHONE, FILE1_10.Outlet,                          FILE6_20H.DOC_DISC, FILE6_20H.CODE AS CUST, FILE6_20H.ISRET, FILE6_20H.ONLINE , SHIP , (FILE6_20H.cash + FILE6_20H.visa + FILE6_20H.point + FILE6_20H.late )  as T_BOON   , FILE6_20.MAN2 " & _
'                    " FROM    FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
'        createCommand cString, con
'
'        cString = " CREATE VIEW [dbo].[STORE_ALL] AS SELECT        CODE, DESCA, BRANCH AS S_BRANCH , ISSTOP FROM            FILE0_40"
'        createCommand cString, con
'
'        cString = " alter TABLE [dbo].[ADDRESS] add   [back] [nvarchar](6) NULL     "
'        createCommand cString, con
'
'    End If
'
'    cString = " alter TABLE [dbo].[FILE6_20H]  add  [visa3] [decimal](18, 2) NOT NULL CONSTRAINT [DF_FILE6_20H_visa3]  DEFAULT ((0)),  [CODEVISA3] [nvarchar](6) NULL,     [TYPEVISA3] [nvarchar](3) NULL,     [VISADOC3] [nvarchar](50) NULL, [visa3_2] [decimal](18, 2) NOT NULL CONSTRAINT [DF_FILE6_20H_visa3_2]  DEFAULT ((0))  "
'    createCommand cString, con
'
''    cString = ReadFile(App.Path & "\TXT\QQFILE6_20.TXT")
''    createCommand cString, con
''
''    cString = ReadFile(App.Path & "\TXT\QQFILE6_20H.TXT")
''    createCommand cString, con
''
''    cString = ReadFile(App.Path & "\TXT\QSALES_MODEL.TXT")
''    createCommand cString, con
''
''    cString = ReadFile(App.Path & "\TXT\QSALESDTL.TXT")
''    createCommand cString, con
''
''    cString = ReadFile(App.Path & "\TXT\QT_SALESDOC.TXT")
''    createCommand cString, con

    
'    cString = "ALTER TABLE [dbo].[FILE0_00] ADD    [OKAZ_ITEM2] [bit] NULL"
'    createCommand cString, con
'    cString = "ALTER TABLE [dbo].[FILE0_00] DROP  COLUMN [OKAZ_ITEM]"
'    createCommand cString, con
'
closeCon con
Err.Clear
End Sub
Private Sub CheckData()
If lServerOnLine Then Exit Sub
If LCase(App.EXEName) <> LCase("JENUR_SHOP_TEMP") Then Exit Sub

On Error Resume Next
con.Execute "CREATE SCHEMA [rp]"
Dim success As Long
Dim sb As New ChilkatStringBuilder
success = sb.LoadFile(App.Path & "\fix.sql", "utf-8")
sMarker = "GO"
Do Until sb.length = 0
    con.Execute sb.GetBefore(sMarker, True)
Loop
Err.Clear
End Sub

