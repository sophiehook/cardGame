# Importing openpyxl to input the data from the workbook in Deck class 
import openpyxl

# Define class Card
class Card:

    shinyNo = 0 

    # Define the initialiser  
    def __init__(self, theName, theType, theHP, theMoves, isShiny):

        """
        __init__ initilises the attributes for the all card objects
        Parameters:
            string - the name of the card
            string - the type the card belongs to 
            int - the health points of the card
            dictionary - the moves as a collection of str names and in damage values
            boolean - isShiny in te representation of 1 or 0 
        Returns:
            Nothing
        """

        # Checks that theName of the card is not empty and of a string data type
        # If it is then theName value is assigned to the card atrribute name
        try:
            if theName is not None:
                if type(theName) == str:
                    self.name = theName
                else:
                    raise TypeError
            else: 
                raise TypeError
        # Throws a type error which is caught by letting the user know there was an issue with the name information
        except:
            print('There was an issue with the name!')

        # Checks that theType of the card is not empty and exists in the array of predefined string types
        # If it is then theType value is assigned to the card atrribute type
        try:
            if theType is not None:
                if theType in ['Magi', 'Water', 'Fire', 'Earth', 'Air','Astral']:
                    self.type = theType
                else: 
                    raise TypeError
            else: 
                raise TypeError
        # Throws a type error which is caught by letting the user know there was an issue with the type information
        except:
            print('There was an issue with the type!')


        # Checks that theHP of the card is not empty and is a postive int 
        # If it is then theHP value is assigned to the card atrribute type
        try:
            if theHP is not None:
                if type(theHP) == int and theHP > 0:
                    self.HP = theHP
                else:
                    TypeError
            else: 
                TypeError
        # Throws a type error which is caught by letting the user know there was an issue with the HP information
        except:
            print('There was an issue with the HP!')


        # Checks that theMoves dictionary of the card is not empty and consists of str values for the move names in the key and postive ints for damage in the corresponding value
        # If it is then theMoves disctionary is assigned to the card atrribute moves
        # Similarly the average damage for each card cannot be asisgned to the cardAverage attrbute unless these is value moves infomation
        try:
            for k,v in theMoves.items():
                if type(k) == str:
                    if type(v) == int and v > 0:
                        self.moves = theMoves
                        self.cardAverage = self.getAverageCard(theMoves)
                    else:
                        raise TypeError
                else:
                    raise TypeError
        # Throws a type error which is caught by letting the user know there was an issue with the moves information
        except:
            print('There was an issue with the moves!')


        # Checks that isShiny is not empty and consists of 0 or 1 corresponding to true and flase
        # If it is then isShiny is assigned to the card atrribute shiny
        try:   
            if isShiny is not None:
                if isShiny == 1:
                    self.shinyNo += 1
                    print(self.shinyNo)
                    self.shiny = isShiny
                elif isShiny == 0:
                    self.shiny = isShiny
                else: 
                    raise TypeError
            else: 
                raise TypeError
        # Throws a type error which is caught by letting the user know there was an issue with the shiny information
        except:
            print('There was an issue with the shiny status!')

        
    # Define the string representation of the card object
    def __str__(self):

        """
        __str__ returns the string representation for card objects
        Parameters:
            None
        Returns:
            string - info on the card
        """

        return 'Name: {self.name}\nType: {self.type}\nHP: {self.HP}\nMoves: {self.moves}\nShiny Status: {self.shiny}'.format(self=self)


    # Gets average damage score for a card 
    def getAverageCard(self, theMoves):

        """
        getAverageCard calculates the damage of the moves in a card 
        Parameters:
            dictionary - the moves as a collection of str names and in damage values
        Returns:
            float - result of the calculation 
        """

        # damageValues holds the numeric values of the dictionary to allow for the sum and len functions
        damageValues = theMoves.values()
        return sum(damageValues) / len(damageValues)
        

# Define class Deck
class Deck:

    # Define the initialiser which creates the card list for each deck
    def __init__(self):

        """
        __init__ initilises the attributes for the all Deck objects
        Parameters:
            Nothing
        Returns:
            Nothing
        """

        self.cards = []


    # Assigns info from the workbook to the variables 
    def inputFromFile(self, fileName):
                
        """
        inputFromFile extracts the information from the cell and assigns it to the correspoding variable 
        Parameters:
            string - the fileName which the user can specify 
        Returns:
            Nothing
        """

        # Trys to find and extract the information from the specified file name
        try:
            book = openpyxl.load_workbook(fileName)
            sheet = book.active

            rowNumber = 0
            for row in sheet.rows:
                # Continues past coloumn names on row 0
                if rowNumber == 0:
                    rowNumber += 1
                    continue

                # Continues past empty rows
                if row[0].value is None:
                    continue

                # Assigns row position vales to variables 
                else:
                    theName = (row[0].value)
                    theType = (row[1].value)
                    theHP = (row[2].value)
                    isShiny = (row[3].value)

                    # Create and add the moves to a dictionary
                    theMoves = {}
                    for i in range(4, (len(row)-1),2):
                        # Continues past empty cells
                        if row[i].value is None:
                            continue
                        # Assigns the name of the move to the key and assigns the corresponding value of damage
                        else: 
                            theMoves[str(row[i].value)] = (row[i+1].value)

                # Adds the newly created Card object with its attributes to the deck in addCard()             
                self.addCard((Card(theName, theType, theHP, theMoves, isShiny)))

                rowNumber += 1

        # Catches exceptions with accessing the file         
        except:
            print("There was an issue with accessing the file," ,fileName , "!")

    # Define the string representation of the Deck object
    def __str__(self):
        
        """
        __str__ returns the string representation for Deck objects
        Parameters:
            None
        Returns:
            string - info on the deck
        """

        return 'Number of cards in deck: '+ str(len(self.cards)) +'\nNumber of shiny cards in deck: '+ str(Card.shinyNo) +'\nAverage damage of the Deck: '+ str(self.getAverageDamage())


    # Adds card objects to the deck list, cards
    def addCard(self, theCard):

        """
        addCard adds the card to the deck
        Parameters:
            Object - the card object 
        Returns:
            Nothing
        """

        # Checks that the object was assigned each attribute in __init__ for the Card class 
        # If it has all these attributes then it is added to the deck list, cards
        try:
            if hasattr(theCard, "name"):
                if hasattr(theCard, "type"):
                    if hasattr(theCard, "HP"):
                        if hasattr(theCard, "moves"):
                            if hasattr(theCard, "shiny"):
                                self.cards.append(theCard)
                            else: 
                                raise TypeError
                        else:
                            raise TypeError
                    else:
                        raise TypeError
                else:
                    raise TypeError
            else:
                raise TypeError

        # If a Card is missing an object it throws an exception
        except:
            print('This card is not valid so will not be added to the deck!')

    # Removes a card from the Deck
    def rmCard(self, theCard):

        """
        rmCard removes a card object from the deck
        Parameters:
            Object - the card object 
        Returns:
            Nothing
        """

        # A card object must be in the deck in the first place in order to remove it 
        try:
            if theCard in self.cards:
                self.cards.remove(theCard)
            else:
                raise NameError

        # Throws an exception if the card doesn't exist  
        except:
            ("You cannot remove that card as it does not exist in the deck!")


    # Finds the most powerful card object in the deck based off it's average damage
    def getMostPowerful(self):
        
        """
        gets the most powerful card from the deck 
        Parameters:
            None
        Returns:
            The most powerful card object in the deck
        """

        highestAvg = 0
        highestAvgIndex = 0

        # Stores the highest damage average index by comparing the current highest average to every other average in the deck
        for card in self.cards:
            if card.cardAverage > highestAvg:

                highestAvg = card.cardAverage
                highestAvgIndex = self.cards.index(card)

        return self.cards[highestAvgIndex]
            
    
    # Calculates the average damage across the whole deck
    def getAverageDamage(self):

        """
        totals the average card damages to have total deck average and divides by number of cards
        Parameters:
            None
        Returns:
            The most powerful card object in the deck
        """

        totalDeckDamage = 0

        for card in self.cards:
            totalDeckDamage += card.cardAverage

        deckAverage = totalDeckDamage/len(self.cards)
        print('Average card damage:',"% .1f" % deckAverage)
        
        return "% .1f" % deckAverage


    # Print the str representation of each card in the deck 
    def viewAllCards(self):

        """
        prints all the cards in the deck
        Parameters:
            None
        Returns:
            Nothing
        """

        # Checks the length of the deck to avoid printing an empty list
        try:
            if len(self.cards)>0: 
                for card in self.cards:
                    print(card)
            else:
                raise NameError

        # Throws and exception 
        except: 
            print("There are no cards in the deck!")    


    # Print the str representation of each shiny card in the deck 
    def viewAllShinyCards(self):

        """
        prints all the shiny cards in the deck
        Parameters:
            None
        Returns:
            Nothing
        """
        try:
            if Card.shinyNo > 0:
                for card in self.cards:
                    if card.shiny == True:
                        print('Card Name:',card.name, '\nShiny status: ', card.shiny)

            # Checks that there is shiny card in the deck 
            else:
                raise TypeError

        # Throws an exception if there are no shiny cards
        except:
           print("There are no shiny cards in this deck!")


    # Prints all the cards of a specified type
    def viewAllByType(self,theType):
                
        """
        Prints all cards in the deck of a certain type
        Parameters:
            None
        Returns:
            Nothing
        """

        # Checks that the type of card is in the deck
        try: 
            for card in self.cards:
                if card.type == theType:
                    print(card)
                else:
                    NameError

        # Throws an exception rather than printing nothing     
        except:
            print("That card type does not exist in the deck")


    # Returns the deck
    def getCards(self):

        """
        gets all the cards of the deck 
        Parameters:
            None
        Returns:
            list - deck of card objects
        """

        return (self.cards)


    # Saves any changes made to the deck in a new file
    def saveToFile(self,fileName): 

        """
        Saves any information stored for the card objects in the deck to a file
        Parameters:
            string - new file name 
        Returns:
            Nothing
        """

        # Trys to append the lists to the workbooks
        try: 
            newBook = openpyxl.Workbook()
            newSheet = newBook.active

            # Column names of the card data and appends those 
            colNames = ('Name','Type','HP','Shiny','Move Name 1', 'Damage 1','Move Name 2', 'Damage 2','Move Name 3', 'Damage 3','Move Name 4', 'Damage 4','Move Name 5', 'Damage 5')
            newSheet.append(colNames)

            # Make a list of all the cards which includes deconstructing the dictionary move and damage pairs to an ordered list 
            for card in self.cards:
                cardList = []
                cardList.append(card.name)
                cardList.append(card.type)
                cardList.append(card.HP)
                cardList.append(card.shiny)

                for k,v in card.moves.items():
                    cardList.append(k)
                    cardList.append(v)

                # Append the card list to the sheet
                newSheet.append(cardList) 
            
            # Save the sheet with the specified file name
            newBook.save(fileName)
        
        # Throws and execption if the file didn't save 
        except: 
            print("Couldn't save to the file", fileName, "!")


    
cd = Deck()
cd.inputFromFile('sampleDeck.xlsx')


