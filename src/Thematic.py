import re
import logging

def check_brackets_balance(rawThematic):
    stack = []
    for char in rawThematic:
        if char == '(':
            stack.append(char)
        elif char == ')':
            if not stack:
                return False
            stack.pop()
    return not stack


def concatenate_thematics(s):
    thematic_list = []
    start = s.find("(")
    while start >= 0:
        end = s.find(")", start)
        thematic = s[start + 1:end].strip()
        logging.info("thematic1=", thematic)
        logging.info("s[end+1:] = ", s[end + 1:])
        if "AND" in s[end + 1:]:
            thematic = "|".join(thematic.split("AND"))
            logging.info("thematic=", thematic)
        thematic_list.append(thematic)
        start = s.find("(", end)
    logging.info("Final List = ", thematic_list)

    return " ".join([s.replace(f"({t})", t) for t in thematic_list if t.strip() != ""])


def find_closing_parenthesis(expression, start_index):
    open_parenthesis_count = 1
    close_parenthesis_count = 0

    for i in range(start_index + 1, len(expression)):
        if expression[i] == '(':
            open_parenthesis_count += 1
        elif expression[i] == ')':
            close_parenthesis_count += 1

        if open_parenthesis_count == close_parenthesis_count:
            return i

    # If no closing parenthesis is found, raise an exception
    raise Exception(f"No closing parenthesis found for expression starting at index {start_index}")


def fetchClosestPair(stack):
    pair = []
    tupPair = []
    for index, tup in enumerate(stack):
        if tup[0] == "open":
            if stack[index + 1][0] == "close":
                tupPair.append(tup[1])
                tupPair.append(index)
                tupPair.append(stack[index + 1][1])
                tupPair.append(index + 1)
                pair.append(tuple(tupPair))
                break
    # logging.info(pair)

    return tuple(tupPair)


def findBracketPair(rawThem):
    stack = []
    pair = []
    openBracketCount = 0
    for index, char in enumerate(rawThem):
        if char == '(':
            stack.append(("open", index))
            openBracketCount = openBracketCount + 1
        if char == ')':
            stack.append(("close", index))
    # logging.info("openBracketCount", openBracketCount)
    # logging.info("Stack Before Extraction = ", stack)
    for i in range(openBracketCount):
        removeItemFromStack = fetchClosestPair(stack)
        tup = stack[removeItemFromStack[1]][1], stack[removeItemFromStack[3]][1]
        pair.append(tup)
        stack.pop(removeItemFromStack[1])
        stack.pop(removeItemFromStack[3] - 1)
        # logging.info("Updated Pair = ", pair)
    return pair


parentFound = []


def findTopParent(bracketPairList):
    parent = []
    child = []
    for bracketPair in bracketPairList:
        # if bracketPair not in parentFound:
        for bp in bracketPairList:
            if bracketPair[0] < bp[0] < bracketPair[1]:
                # logging.info(f"Parent -> {bracketPair} , Child -> {bp}")
                child.append(bp)
                parent.append(bracketPair)
        count = 0
        for bp in bracketPairList:
            if bracketPair[0] < bp[0] and bracketPair[1] > bp[1]:
                count = count + 1
        if count == 0:
            # logging.info(f"{bracketPair} is a parent and has no children")
            parent.append(bracketPair)
    # logging.info("parent = ", parent)
    # logging.info("child = ", child)
    finalParentList = []
    # Remove child from parent list
    for item in parent:
        if item not in child:
            finalParentList.append(item)
    finalParentList = [*set(finalParentList)]
    # logging.info("finalParentList = ", finalParentList)
    return finalParentList


# returns a list of tuples contains index position of redundant brackets, which can be removed.
def removeRedundantBrackets(pairList):
    removeIndex = []
    # logging.info(pairList)
    for index, item in enumerate(pairList):
        for i in pairList:
            if item[0] == i[0] + 1 and item[1] == i[1] - 1:
                removeIndex.append(item[0])
                removeIndex.append(item[1])

    return removeIndex


def preProcess(them):
    return them


# Takes parent tuple as input and returns corresponding children as list of tuple
def getChild(pairList, tupParent):
    children = []
    for pair in pairList:
        if pair[0] > tupParent[0] and pair[1] < tupParent[1]:
            children.append(pair)
    # remove child of child
    indexToRemove = []
    for child in children:
        for index, ch in enumerate(children):
            if child[0] < ch[0] and child[1] > ch[1]:
                indexToRemove.append(index)
    # logging.info(f" Final child list ", [child for index, child in enumerate(children) if index not in indexToRemove])
    return [child for index, child in enumerate(children) if index not in indexToRemove]


# Gets index (as a tuple) of two elements  and returns the relation
def getRelation(thematic, elements):
    relation = []
    for index, element in enumerate(elements):
        if index < len(elements) - 1:
            if "AND" in thematic[elements[index][1]:elements[index + 1][0]]:
                relation.append("AND")
            elif "OR" in thematic[elements[index][1]:elements[index + 1][0]]:
                relation.append("OR")
            else:
                relation.append("AND")
    return relation


def getAllThematicCode(thematic, element):
    pattern1 = r"[a-zA-Z0-9]{3}_[0-9]{2}"
    pattern2 = r"([a-zA-Z0-9]{3}_[0-9]{2} *AND *[a-zA-Z0-9]{3}_[0-9]{2})+|( *AND *[a-zA-Z0-9]{3}_[0-9]{2}) * *\)+$"
    pattern3 = r'(?:[A-Za-z0-9]{3}_\d{2}\s+AND\s+)*[A-Za-z0-9]{3}_\d{2}'
    thematic = re.findall(pattern3, thematic[element[0]:element[1]])
    return thematic


def combineTwoNodes(node1, node2):
    themLine = ""
    return themLine


# Function to sort the list by first item of tuple
def Sort_Tuple(tup):
    tup.sort(key=lambda x: x[0])
    return tup


def convertRawToThematicLines(rawThematics, bracketlogging=None):
    thematicLine = []
    # logging.info("rawThematics = ", rawThematics)
    processedRawThem = preProcess(rawThematics)
    isBracketBalanced = check_brackets_balance(processedRawThem)
    bracketlogging.infoMsg = "Brackets are Balanced" if isBracketBalanced else "Brackets are not balanced. Correct the " \
                                                                        "Brackets and rerun the tool. "

    updatedPair = []
    if isBracketBalanced:
        updatedPair = findBracketPair(processedRawThem)
        # Remove redundant
        indexToRemove = removeRedundantBrackets(updatedPair)
        themNoRedundant = ""
        for index, char in enumerate(processedRawThem):
            if index not in indexToRemove:
                themNoRedundant = themNoRedundant + char
        updatedPair = findBracketPair(themNoRedundant)

        # find top parents
        topParents = findTopParent(updatedPair)

        # Arrange the elements in ascending
        topParents = Sort_Tuple(topParents)


        if len(topParents) == 1:
            themNoRedundant = themNoRedundant[topParents[0][0]+1:topParents[0][1]]
            updatedPair = findBracketPair(themNoRedundant)
            topParents = findTopParent(updatedPair)
            topParents = Sort_Tuple(topParents)


        # logging.info("Updated Pair = ", topParents)
        # Get Relation
        relationList = getRelation(themNoRedundant, topParents)

        # GetAllThematics inside element
        topParentsThem = []
        for element in topParents:
            # logging.info("Element = ", element)
            topParentsThem.append(getAllThematicCode(themNoRedundant, element))

        # arrange the list which contains less element first if all element of relationList is AND
        if all(element == 'AND' for element in relationList):
            topParentsThem = sorted(topParentsThem, key=len)

        for index, elements in enumerate(topParentsThem):
            for element in elements:
                if index + 1 < len(topParentsThem):
                    for i, next in enumerate(topParentsThem[index + 1]):
                        if relationList[index] == "AND":
                            topParentsThem[index + 1][i] = element + " AND " + next
                        else:
                            ...
        finalThematicLine = []
        if all(element == 'OR' for element in relationList):
            for thematicLine in topParentsThem:
                for them in thematicLine:
                    logging.info(f"them {them}...............\n")
                    # global final
                    final = them.replace("AND", "|")
                    final = final.replace(" ", "")
                    finalThematicLine.append(final)
        elif all(element == 'AND' for element in relationList):
            for thematicLine in topParentsThem[len(topParentsThem) - 1]:
                # global final
                logging.info(f"\n\nthematicLine {thematicLine}...............\n")
                final = thematicLine.replace("AND", "|")
                final = final.replace(" ", "")
                finalThematicLine.append(final)
        else:
            # logging.info("Node contains both AND , OR relation - ???")
            return -2
        return finalThematicLine
    else:
        logging.info(bracketlogging.infoMsg)
        return -1

        # combineTwoNodes()

        # Has Parent

        # Has Child

        # If Node has no child, Internally solve AND

        # Check if the node contains OR, if yes, process OR


if __name__ == "__main__":
    # rawThematics = " (((JAB_01 OR HGV_01 OR GDF_02 OR SDF_02)) AND (ASD_01 AND DFK_01 AND DFG_09 AND ASH_02) AND (GHF_03) AND (DFG_04))"
    # rawThematics = "((ALO_01 AND ALO_02 AND ALO_03) AND (DFH_05))"
    # rawThematics = "AND (DFH_05 AND ALO_01) OR (DFH_05 AND ALO_02) OR (DFH_05 AND ALO_03)"
    # rawThematics = "   AND ((ALO_01 AND ALO_02 AND ALO_03) AND (DFH_05))"
    # rawThematics = "(ALO_01 AND ALO_02 AND ALO_03) AND (DFH_05)"
    # rawThematics = " AND (ALO_02) AND (DFH_05) AND (LLL_01)"
    # rawThematics = " ((( JAB_01 ))) AND (( CYN_01 ) )"
    # rawThematics = " (JAB_01) AND (CYN_01 AND  LLL_01  OR  DL8_01  AND  LLL_02  )"
    # rawThematics = " (JAB_01)  AND ( CYN_01  AND  LLL_01  OR  DL8_01  AND  LLL_02 AND ABH_02 OR FGG_09 AND ABV_01 ) AND (ASD_001) AND (HYK_01)"
    # rawThematics = " (JAB_01) AND ( ( CYN_01 OR (ABC_01 AND VGH_01 OR (AMN_02)) ) AND ( LLL_01 ) OR ( DL8_01 AND (SXC_01)) AND ( LLL_02 ) )"
    # rawThematics = "(JAB_01)AND((CYN_01 AND (ABC_01 OR (XYZ_01 (CBV_01 ))))AND(LLL_01)OR(DL8_01)AND(LLL_02))"
    # rawThematics = " ( JAB_01 OR AMC_01) AND (( CYN_01 ) AND ( LLL_01 ) OR ( DL8_01 ) AND ( LLL_02 ) ) AND ( (XYZ_01) AND (ABC_01) )"
    # rawThematics = " ( ( JAB_01 ) AND (( CYN_01 )) AND ( LLL_01 ) OR ( DL8_01 ) AND ( LLL_02 ) )"
    # rawThematics = " ( (  ( JAB_01 )  ) AND ( ( ( CYN_01 )  ) AND (  ( LLL_01 ) ) OR (  ( DL8_01 )  ) AND ( ( LLL_02 )  ) ) )"
    # output = convert_expression(expression)
    # rawThematics = " (JAB_01) "
    # rawThematics = " ((( JAB_01 ))) AND (( CYN_01 ) )"
    # rawThematics = " ( DDC_25 AND DUW_01 AND ACB_01 ) "
    # rawThematics = "  (DDC_25 OR DUW_01 ) AND ( ACB_01    AND GHG_02 AND KLH_02) AND (GJG_02 AND KLK_09 AND JKJ_07)"
    # rawThematics = "  ( AFC_02 AND AFE_04 ) OR ( AFC_02 AND AFE_06 ) OR ( AFE_04 AND AFC_04 ) OR ( AFE_06 AND AFC_04 ) "
    # rawThematics = "   ( LYQ_01 AND AFC_02 AND BFC_01 ) OR ( LYQ_02 AND AFC_02 AND DXD_00 ) OR ( AFC_04 ) "
    # rawThematics = "   ( LYQ_01 AND AFC_02 AND BFC_01 AND IOP_07 ) OR ( AFC_04 AND IOP_07 ) OR ( LYQ_02 AND AFC_02 AND DXD_00 AND IOP_07 )  "
    # rawThematics = " ( ( LYQ_01 AND AFE_06 AND IOC_03 ) ) OR ( ( LYQ_01 AND AFE_06 AND IOC_02 ) ) OR ( ( LYQ_01 AND LYQ_02 AND AFE_06 AND D7T_02 ) )  "
    # rawThematics = "( BFC_02 AND LYQ_01 AND AFC_02 ) OR ( LYQ_01 AND AFC_02 AND BFC_03 ) OR ( LYQ_02 AND DXD_03 AND AFC_02 ) OR ( DXD_04 AND AFC_02 AND LYQ_02 ) OR ( LYQ_02 AND AFC_02 AND DXD_05 ) OR ( AFC_02 )  "
    # rawThematics = "( LYQ_02 AND AFC_02 AND DXD_00 AND AFE_02 AND IOP_07 ) OR ( LYQ_02 AND AFE_03 AND IOP_07 AND AFC_02 AND DXD_00 ) OR ( LYQ_02 AND AFE_05 AND IOP_07 AND AFC_02 AND DXD_00 ) OR ( IOP_07 AND LYQ_01 AND AFE_01 AND BFC_01 AND AFC_02 ) OR ( BFC_01 AND AFE_02 AND LYQ_01 AND AFC_02 AND IOP_07 ) OR ( AFE_03 AND IOP_07 AND AFC_02 AND BFC_01 AND LYQ_01 ) OR ( BFC_01 AND AFE_05 AND AFC_02 AND LYQ_01 AND IOP_07 ) "
    # rawThematics = "( LYQ_01 AND AEN_03 AND AFC_02 AND BFC_01 ) OR ( BFC_01 AND LYQ_01 AND AFC_02 AND AEN_01 ) OR ( DVQ_52 AND LYQ_02 AND DXD_00 AND AFC_02 ) OR ( DXD_00 AND DVQ_54 AND LYQ_02 AND AFC_02 ) OR ( DXD_00 AND AFC_02 AND LYQ_02 AND DVQ_61 ) OR ( AFC_02 AND DXD_00 AND LYQ_02 AND DVQ_62 )   "
    # rawThematics = " ( DUW_01 AND DDC_25 OR SDF_01 AND DFG_01) OR ( DUW_01 AND DDC_32 ) OR ( DUW_01 AND DDC_32 ) OR ( DUW_02 AND DDC_01 )  "
    # rawThematics = "( ( LLL_02 AND DAO_01 AND LNG_02 ) ) OR ( ( LLL_02 AND DAO_01 AND LNG_03 ) ) OR ( ( LLL_02 AND DAO_03 ) ) "
    # rawThematics = "( ( XYZ_02 OR ABC_01 OR YUI_02 OR ASC_01) ) OR ( ( LLL_02 AND DAO_01 AND LNG_03 ) ) OR ( ( LLL_02 AND DAO_03 ) )  OR (DFG_01 OR FGH_02)"
    # rawThematics = "(ALO_01 AND ALO_02) AND (DFH_05)"
    # rawThematics = "AND (DFH_05 AND LYQ_01)"
    # rawThematics = "AND ((DFH_05) AND (ALO_01)) OR ((DFH_05) AND (ALO_02))"
    # rawThematics = "AND (EEK_00) AND (IZB_00) AND (LUG_01) AND (LWK_02) AND (EEN_00 AND IZE_05 OR IZE_06 OR IZE_07 OR IZE_08) AND (LME_00 AND IZA_00 OR JXN_00 AND AWM_03)"
    # rawThematics = "AND (JAB_01) AND (CYN_01 AND LLL_01 OR DL8_01 AND LLL_02)"
    rawThematics = "AND ((DFH_05) AND (LYQ_01))"
    # rawThematics = "AND ( CLI TYPE_AEE_LEVE_VITRES (CLI_02 MUX)  AND IWV OPTION_LV_AR_ELEC (IWV_00 WITHOUT)  AND IWY OPTION_LV_AP (IWY_01 AVEC_AP)  AND LNG TYPE_SIDE_DOORS_ARCHI (LNG_02 2_DCU_AV , LNG_03 4_DCU_AV_AR)   AND LYQ TYPE_DIVERSITY (LYQ_01 BEFORE_FUNCT_CODIF)  )  OR ( CLI TYPE_AEE_LEVE_VITRES (CLI_02 MUX)   AND DLE REAR WINDOWS LIFTER (DLE_00 WITHOUT , DLE_10 MANUAL)  AND IWY OPTION_LV_AP (IWY_01 AVEC_AP)  AND LNG TYPE_SIDE_DOORS_ARCHI (LNG_02 2_DCU_AV , LNG_03 4_DCU_AV_AR)  AND LYQ TYPE_DIVERSITY (LYQ_02 FUNCT_CODIF)  ) "
    logging.info("=" * 100)
    logging.info("Raw Thematic = ", rawThematics)
    thematicLines = convertRawToThematicLines(rawThematics)
    print("thematicLines-------------->", thematicLines)
    for thematicLine in thematicLines:
        if thematicLine != -1 or thematicLine != -2:
            logging.info(thematicLine)
    logging.info("=" * 100)

# Limitation thematic = ((ALO_01 AND ALO_02) AND (DFH_05))
# ((ALO_01) AND (ALO_02) AND (HGB_03))
# (ALO_01 AND ALO_02) AND (DFH_05) OR (JFJ_02)
