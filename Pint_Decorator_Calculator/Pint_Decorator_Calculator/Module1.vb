Module Module1
    'Created By Mohammed Shakil Uddin
    'Date Started 28/02/14
    'Paint Decorator Calculator
    'Final Version 23/03/14

    'These are all my Global Variables

    'this is my variable for my options procedure

    Dim Options As Integer ' this is used for my Menu_option

    'these are my variable for the room measurements 

    Dim WidthSize As Single ' this is for my estimate paint cost for the Width
    Dim LengthSize As Single ' this is for my estimate paint cost for the Length
    Dim HeightSize As Single ' this is for my estimate paint cost for the Height

    'this is an added section which will allow the user to change the measurements of the room if they wish to do so

    Dim ErrorInRoomSize As String ' this will help my clients change her answer if she made any mistakes
    Dim ChangeAnswer As String ' This will be used to change the answer (reason why I have three is because Width, Length and Height are the 3 measurements)

    'these are my variables which is used for the paint quality for my user to choose and the cost and value too 

    Dim PaintChoice As String ' this feature will allow the user to choose what paint they want to use also if they make mistake, it will repeat that part again
    Dim PaintQuality As String ' this is where the user will choose what paint quality they want to use
    Dim PaintPrice As Single ' this is the price of the paint quality

    'these variables are used for the undercoating that my user can use if they wish to do so

    Dim UnderCoat As String ' This is the UnderCoat where user will choose
    Dim UnderCoatPrice As Single = 0.5 ' This is the UnderCoat Price

    'these variables are used for the calculations for the area of the room, the total price with and without undercoat
    'also these calculations show the overall price with the total price with and without VAT 

    Dim AreaOfRoom As Single ' This is the calculation of the area
    Dim TotalPrice As Single ' This is the calculation of the area
    Dim TotalPriceUnderCoat As Single ' This is the calculation of the area
    Dim TotalPriceWithUnderCoatAndVAT As Single ' This is the calculation of the area
    Dim JustVATPrice As Single ' This is the calculation of the area
    Dim TotalPriceWithUnderCoatWithoutVAT As Single ' This is the calculation of the area
    Sub Main() ' this will be calling the title and the menu options together in that order
        Call Title() ' this is a function that will call the title procedure which will show the name of the program
        Call Menu_Option() ' this function will call the menu option that will now give the user the options to go to one part of the program depending what the user chooses
    End Sub

    Sub Title() 'This procedure is my title
        Console.BackgroundColor = ConsoleColor.Yellow ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Red ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("|==============================================================================|")
        Console.BackgroundColor = ConsoleColor.White ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.DarkGray ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("|====================Welcome to my Paint Decorator Calculator==================|")
        Console.BackgroundColor = ConsoleColor.Red ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Black ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("|==============================================================================|")
    End Sub
    Sub Menu_Option() ' this procedure is when user will choose an options
        Console.BackgroundColor = ConsoleColor.DarkRed ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Yellow ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("Please select a number from the options below")
        Console.BackgroundColor = ConsoleColor.White ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Black ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("1.Estimated Paint Cost")
        Console.WriteLine("2.View Estimate Total")
        Console.WriteLine("3.FAQ & Help")
        Console.WriteLine("4.Exit Program")
        Console.BackgroundColor = ConsoleColor.Black ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.BackgroundColor = ConsoleColor.White ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Try ' this piece of coding will allow this procedure to repeat again if the user puts in the wrong data type in as a value
            Console.WriteLine("please enter a number between 1 - 4")
            Options = Console.ReadLine ' this is the assignment which will carry the same value whatever the user types in
        Catch ex As Exception
            Console.WriteLine("Error, please enter a number between 1 - 4")
            Console.WriteLine("Please press on enter to start again")
            Console.BackgroundColor = ConsoleColor.Black ' this piece of coding will change the colour of the background depending on the colour I have chosen
            Console.ReadLine()
            Console.Clear()
            Call Menu_Option()
        End Try
        If Options = 1 Then ' these are my IF Statement that will take the user to different procedure depending on the input they give
            Call Estimated_Paint_Cost() ' this will call the procedure called Estimated_Paint_Cost
        ElseIf Options = 2 Then
            Call View_Estimate_Total() ' this will call the procedure called Estimate_Total
        ElseIf Options = 3 Then
            FAQ_And_Help() ' this will call the procedure called FAG_And_Help
        ElseIf Options = 4 Then
            Exit_Page() ' this will call the procedure to close the program
            Call Exit_Page()
        Else ' this procedure will repeat this procedure again if user puts in an invalid letter
            Console.WriteLine("Error please enter between 1 - 4")
            Menu_Option()
        End If ' this is where the if statement finishes for this section
    End Sub
    Sub Estimated_Paint_Cost() ' this procedure is my first option which will now ask the user to put in values for each of these section
        Console.BackgroundColor = ConsoleColor.Black
        Console.Clear() ' this piece of coding will clear out all the text before it calls in the following procedures
        Call Width_Size() ' this will call the Width size procedure which will ask for the user to type in the users measurements
        Call Length_Size() ' this will call the Length size procedure which will ask for the user to type in the users measurements
        Call Height_Size() ' this will call the Height size procedure which will ask for the user to type in the users measurements
        Call Change_If_Mistake() ' this is an option that will allow the user to repeat the measurements again if they were to type in their room measurements incorrectly
        Call Change_Answer() ' this is a small procedure which will call the Width_Size so the user can type in the measurements again
        Call Paint_Quality() ' this will call the paint quality which will allow the user to choose the paint quality
        Call Under_Coat() ' this will call the undercoat which allow the user to decide if they want the under coating
        Call Total_Price_Without_UnderCoat() ' this procedure will show the amount the user will need to pay after they have answered all the questions without undercoat
        Call Total_Price_With_UnderCoat() ' this procedure will show the amount the user will need to pay after they have answered all the questions with undercoat
    End Sub
    Sub Width_Size()
        Console.BackgroundColor = ConsoleColor.Cyan
        Console.ForegroundColor = ConsoleColor.DarkMagenta
        Try
            Console.WriteLine("Input the Width of the room between 1 to 25: ")
            WidthSize = Console.ReadLine
        Catch ex As Exception
            Console.WriteLine("Error, please input the Width of the room between 1 to 25: ")
            Call Width_Size()
        End Try
   
        If WidthSize < 1 Or WidthSize > 25 Then
            Console.WriteLine("Error, You Must Enter between 1 to 25 ")
            Call Width_Size()
        End If
    End Sub
    Sub Length_Size()
        Console.BackgroundColor = ConsoleColor.Blue
        Console.ForegroundColor = ConsoleColor.Green
        Try
            Console.WriteLine("Input the Length of the room between 1 to 25: ")
            LengthSize = Console.ReadLine
        Catch ex As Exception
            Console.WriteLine("Error, please input the Length of the room between 1 to 25: ")
            Call Length_Size()
        End Try
        If LengthSize < 1 Or LengthSize > 25 Then
            Console.WriteLine("Error, You Must Enter between 1 to 25 ")
            Call Length_Size()
        End If
    End Sub
    Sub Height_Size()
        Console.BackgroundColor = ConsoleColor.Gray ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.DarkGreen ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("Input the Height of the room between 2 to 6: ")
        HeightSize = Console.ReadLine ' this is the assignment which will carry the same value whatever the user types in
    End Sub
    Sub Change_If_Mistake()
        Console.BackgroundColor = ConsoleColor.Red ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Yellow ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Try ' this piece of coding will allow this procedure to repeat again if the user puts in the wrong data type in as a value
            Console.WriteLine("If you have made an Error or wish to change the measurements again type Yes, if not say No [Y/N]: ")
            ErrorInRoomSize = Console.ReadLine ' this is the assignment which will carry the same value whatever the user types in
        Catch ex As Exception
            Console.WriteLine("Error please enter Y or N, Must be in CAPITAL LETTERS")
            Console.WriteLine("Please press on enter to start again")
            Call Change_If_Mistake()
        End Try

        If ErrorInRoomSize = "Y" Then
            Call Change_Answer()
        End If
        If ErrorInRoomSize = "N" Then
            Call Paint_Quality()
        Else ' this procedure will repeat this procedure again if user puts in an invalid letter
            Console.WriteLine("Error please enter Y or N, Must be in CAPITAL LETTERS")
            Change_If_Mistake()
        End If
    End Sub
    Sub Change_Answer()
        Call Estimated_Paint_Cost()
    End Sub
    Sub Paint_Quality() 'this is were the user will choose the paint quality
        Console.BackgroundColor = ConsoleColor.White ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Black ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("a. Luxury Paint Quality ---- £1.75")
        Console.WriteLine("b. Standard Paint Quality ---- £1.00")
        Console.WriteLine("c. Economy Paint Quality ---- £0.45")
        Try ' this piece of coding will allow this procedure to repeat again if the user puts in the wrong data type in as a value
            Console.WriteLine("Please select a letter between a - c")
            PaintChoice = Console.ReadLine ' this is the assignment which will carry the same value whatever the user types in
        Catch ex As Exception
            Console.WriteLine("Error, please enter a number between a - c")
            Console.WriteLine("Please press on enter to start again")
            Call Paint_Quality()
        End Try
        ' these are my IF Statement that will take the user to different procedure depending on the input they give
        If PaintChoice = "a" Then
            PaintChoice = "Luxury Paint"
            PaintPrice = FormatCurrency(1.75)
            Call Under_Coat()
        End If
        If PaintChoice = "b" Then
            PaintChoice = "Standard Paint"
            PaintPrice = FormatCurrency(1.0)
            Call Under_Coat()
        End If
        If PaintChoice = "c" Then
            PaintChoice = "Economy Paint"
            PaintPrice = FormatCurrency(0.45)
            Call Under_Coat()
        Else
            Console.WriteLine("Error please enter between a - c")
            Paint_Quality()
        End If

    End Sub
    Sub Under_Coat()
        Console.BackgroundColor = ConsoleColor.White ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Red ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Try ' this piece of coding will allow this procedure to repeat again if the user puts in the wrong data type in as a value
            Console.WriteLine("Do you wish to buy UnderCoat? [Y/N]")
            UnderCoat = Console.ReadLine ' this is the assignment which will carry the same value whatever the user types in
        Catch ex As Exception
            Console.WriteLine("Error please enter Y or N, Must be in CAPITAL LETTERS")
            Console.WriteLine("Please press on enter to start again")
            Console.ReadLine()
            Call Under_Coat()
        End Try
        ' these are my IF Statement that will take the user to different procedure depending on the input they give
        If UnderCoat = "Y" Then
            UnderCoatPrice = 0.5
            Call Total_Price_With_UnderCoat()
        ElseIf UnderCoat = "N" Then
            UnderCoatPrice = 0
            Call Total_Price_Without_UnderCoat()
        Else ' this procedure will repeat this rule again if they type in an invalid letter
            Console.WriteLine("Error please enter Y or N, Must be in CAPITAL LETTERS")
            Under_Coat()
        End If
        ' this is wHere the if statemen finish for this section
    End Sub
    Sub Total_Price_With_UnderCoat()
        Console.BackgroundColor = ConsoleColor.Cyan
        Console.ForegroundColor = ConsoleColor.DarkCyan
        AreaOfRoom = ((LengthSize * HeightSize) * 2) + ((WidthSize * HeightSize) * 2)
        TotalPrice = PaintPrice * AreaOfRoom
        TotalPriceUnderCoat = UnderCoatPrice * AreaOfRoom

        TotalPriceWithUnderCoatWithoutVAT = TotalPriceUnderCoat + TotalPrice
        JustVATPrice = TotalPriceWithUnderCoatWithoutVAT * 0.2
        TotalPriceWithUnderCoatAndVAT = (TotalPriceUnderCoat + TotalPrice) * 1.2
        Console.WriteLine("The room size is " & AreaOfRoom & " Meters")
        Console.WriteLine("The paint cost is " & FormatCurrency(TotalPrice))
        Console.WriteLine("The price for Undercoat only is " & FormatCurrency(TotalPriceUnderCoat))
        Console.WriteLine("Therefore the overall price without VAT is £" & (TotalPriceWithUnderCoatWithoutVAT))
        Console.WriteLine("The VAT price is " & FormatCurrency(JustVATPrice))
        Console.WriteLine("But the overall price with VAT is " & FormatCurrency(TotalPriceWithUnderCoatAndVAT))
        Console.WriteLine("If you wish to look at the results separately in full details please go to the")
        Console.WriteLine("total price again please go to 2.View Estimate Total")
        Console.BackgroundColor = ConsoleColor.Black
        Console.ReadLine()
        Console.Clear()
        Call Title()
        Call Menu_Option()
    End Sub
    Sub Total_Price_Without_UnderCoat()
        Console.BackgroundColor = ConsoleColor.Cyan
        Console.ForegroundColor = ConsoleColor.DarkCyan
        AreaOfRoom = ((LengthSize * HeightSize) * 2) + ((WidthSize * HeightSize) * 2)
        TotalPrice = PaintPrice * AreaOfRoom
        TotalPriceUnderCoat = UnderCoatPrice * AreaOfRoom

        TotalPriceWithUnderCoatWithoutVAT = TotalPriceUnderCoat + TotalPrice
        JustVATPrice = TotalPriceWithUnderCoatWithoutVAT * 0.2
        TotalPriceWithUnderCoatAndVAT = (TotalPriceUnderCoat + TotalPrice) * 1.2
        Console.WriteLine("The room size is " & AreaOfRoom & " Meters")
        Console.WriteLine("The paint cost is " & FormatCurrency(TotalPrice))
        Console.WriteLine("Therefore the overall price without VAT is " & FormatCurrency(TotalPriceWithUnderCoatWithoutVAT))
        Console.WriteLine("The VAT price is " & FormatCurrency(JustVATPrice))
        Console.WriteLine("But the overall price with VAT is " & FormatCurrency(TotalPriceWithUnderCoatAndVAT))
        Console.WriteLine("If you wish to look at the results separately in full details please go to the")
        Console.WriteLine("total price again please go to 2.View Estimate Total")
        Console.BackgroundColor = ConsoleColor.Black
        Console.ReadLine()
        Console.Clear()
        Call Title()
        Call Menu_Option()
    End Sub
    Sub View_Estimate_Total()
        Console.BackgroundColor = ConsoleColor.Black
        Console.Clear()
        Console.BackgroundColor = ConsoleColor.DarkCyan
        Console.ForegroundColor = ConsoleColor.Yellow
        Console.WriteLine("|------------------------------------------------------------------------------|")
        Console.BackgroundColor = ConsoleColor.Cyan
        Console.ForegroundColor = ConsoleColor.Red
        Console.WriteLine("|--------------------------Here is your Total Price----------------------------|")
        Console.BackgroundColor = ConsoleColor.DarkMagenta
        Console.ForegroundColor = ConsoleColor.Blue
        Console.WriteLine("|------------------------------------------------------------------------------|")
        Console.BackgroundColor = ConsoleColor.White
        Console.ForegroundColor = ConsoleColor.Black
        Console.WriteLine("                                                  |================|")
        Console.WriteLine("The Room aread for all the walls are:             |" & AreaOfRoom & " Square Meters")
        Console.WriteLine("The Paint Price for the quality you chosen us is: |" & PaintChoice)
        Console.WriteLine("Using Undercoat?:                                 |" & UnderCoat)
        Console.WriteLine("Undercoat Price per meter is:                     |" & FormatCurrency(UnderCoatPrice))
        Console.WriteLine("Total Price Without VAT:                          |" & FormatCurrency(TotalPriceWithUnderCoatWithoutVAT))
        Console.WriteLine("VAT Price Is:                                     |" & FormatCurrency(JustVATPrice))
        Console.WriteLine("Total Price With VAT:                             |" & FormatCurrency(TotalPriceWithUnderCoatAndVAT))
        Console.WriteLine("                                                  |================|")
        Console.BackgroundColor = ConsoleColor.Black
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine("Please Enter to start again")
        Console.ReadLine()
        Console.Clear()
        Call Title()
        Call Menu_Option()
    End Sub
    Sub FAQ_And_Help() 'this procedure is the FAQs and Help page which will explain how to use the program
        Console.BackgroundColor = ConsoleColor.Cyan ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.DarkCyan ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("--------------------------------------------------------------------------------")
        Console.BackgroundColor = ConsoleColor.Blue ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Yellow ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("------------------------This is the FAQ & Help Page-----------------------------")
        Console.BackgroundColor = ConsoleColor.Green ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.DarkCyan ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("--------------------------------------------------------------------------------")
        Console.BackgroundColor = ConsoleColor.Yellow ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.DarkMagenta ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("1. Here is a Tutorial on how to use my program.")
        Console.WriteLine("1.1 You will be needing The keyboard for this Application.")
        Console.WriteLine("1.2 In this program you will be given options, the program will give you numbers that you can put in as minimum and maximum.")
        Console.WriteLine("1.3 Depending on the number you will choose the application will take you there")
        Console.WriteLine("1.4 When you are on the estimate paint cost, you will be given questions that")
        Console.WriteLine("you will need fill up with the correct answer.")
        Console.WriteLine("1.5 Different Pages will give different responses")
        Console.WriteLine("1.6 View Estimate Cost will show you how much it will cost you depending what")
        Console.WriteLine("you put in.")
        Console.WriteLine("1.7 View Estimate Total will show you the total it will cost you, also it will")
        Console.WriteLine("show you what you have typed in.")
        Console.BackgroundColor = ConsoleColor.DarkYellow ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Green ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("2. Now I will be explaining how to make the order, from this point you will")
        Console.WriteLine("only need your keyboard to type.")
        Console.WriteLine("2.1 To make the order you will need to go to Estimate Paint Cost and type")
        Console.WriteLine("in the measurements of your room.")
        Console.WriteLine("2.2 After that the program will ask you to Choose the paint quality that you")
        Console.WriteLine("would want to use.")
        Console.WriteLine("2.3 The next step that the program will ask you if you want undercoating")
        Console.WriteLine("which is another layer on top of your paint.")
        Console.WriteLine("2.4 Finally when you have made your decision on the undercoat, it would give")
        Console.WriteLine("you the full details eg room size, paint quality, undercoating, overall price.")
        Console.BackgroundColor = ConsoleColor.DarkBlue ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Red ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("3. You are able to look back at the order you have created")
        Console.WriteLine("3.1 To look back out your order go to the View Estimate Total.")
        Console.WriteLine("3.2 This page will allow you to look back at the order you have chosen.")
        Console.WriteLine("3.3 It will show you the room size, paint price per meter, undercoat if you.")
        Console.WriteLine("have chosen it, the price with and with VAT.")
        Console.BackgroundColor = ConsoleColor.Green ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Magenta ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("4. This page will show you how to exit the program")
        Console.WriteLine("4.1 To exit the program just go to Exit Program in the menu options.")
        Console.BackgroundColor = ConsoleColor.DarkCyan ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.White ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine("Contact Details are:")
        Console.WriteLine("Email:  shakil-uddin@hotmail.co.uk")
        Console.WriteLine("Phone:  0208 517 12 15")
        Console.WriteLine("Mobile: 07572114660")
        Console.BackgroundColor = ConsoleColor.Cyan ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.ForegroundColor = ConsoleColor.Red ' this piece of coding will change the colour of the text depending on the colour I have chosen
        Console.WriteLine()
        Console.WriteLine("|==================== FAQ ==================|")
        Console.WriteLine()
        Console.WriteLine("How to Put in the measurements of my room in this program?")
        Console.WriteLine("To enter the measurements of the room you need to go to Estimate Paint Cost")
        Console.WriteLine()
        Console.WriteLine("How do you use the Undercoat?")
        Console.WriteLine("When you have typed in the measurements of your room then you will")
        Console.WriteLine("be asked if you want undercoat.")
        Console.WriteLine()
        Console.WriteLine("How to get the full results?")
        Console.WriteLine("To get the results you will need to answer the questions asked")
        Console.WriteLine("on the Estimate the Total, after that then go to View Estimate Total")
        Console.ReadLine()
        Console.BackgroundColor = ConsoleColor.Black ' this piece of coding will change the colour of the background depending on the colour I have chosen
        Console.Clear()
        Menu_Option() ' this procedure is when user will choose an options
        Console.WriteLine("please enter a number between 1 - 4")
        Options = Console.ReadLine
    End Sub
    Sub Exit_Page() 'this procedure is the exit page which will close the program
        End
    End Sub

End Module

