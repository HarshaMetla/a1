import rhparserv1 as rh
import re
import camelot
import pandas as pd

def camelot_extraction(filepath,tokens_for_processing):
    tables = camelot.read_pdf(filepath, pages='all')
    #print(tables)
    #print(tokens_for_processing)
    #print(tables[0].df[0][0].split(" ")[0].replace("：", ""))
    indexreq = rh.get_identitytoken(tokens_for_processing,tables[0].df[0][0].split(" ")[0].replace("：", ""))[1]
    # print(indexreq)
    headertokens = rh.get_between_tokens(tokens_for_processing, 0, indexreq)
    # print(headertokens)
    headerdata = " ".join(headertokens)
    #print("header data"+"\n", headerdata)
    string = headerdata
    lasteleindex = 0
    no_of_tables = len(tables)
    for n in range(len(tables)):
        def regex(x):
            lst = [
                re.sub("[⺀-⺙⺛-⻳⼀-⿕々〇〡-〩〸-〺〻㐀-䶵一-鿃豈-鶴侮-頻並-龎]", '', str(i)) for i in x]
            return pd.Series(lst)

        df = tables[n].df
        df1 = df.apply(regex, result_type='expand', axis=1)
        df2 = df1.replace('\n', '', regex=True)
        df3 = df1.replace('\n', ' ', regex=True)
        final = df3.apply(lambda col: col.astype(str)) + " " + df2
        for i in final.columns:
            final[i] = final[i].apply(lambda x: ' '.join(sorted(set(x.split(' ')))))
        final1 = final.to_string(header=False, index=False)
        string = string+" "+final1
        #print("last row", tables[n].df.iloc[-1].to_list())
        cleanedList = [x for x in tables[n].df.iloc[-1] if str(x) != ""]
        # print("cleaned list", cleanedList)
        # print("cleaned last element",cleanedList[-1].split(' ')[-1].replace('。', '').replace('\n', ''))
        for token in rh.get_token_atIndexRange(tokens_for_processing, lasteleindex+1, tokens_for_processing[-1][1]):
            if token[0] == cleanedList[-1].split(' ')[-1].replace('。', '').replace('\n', ''):
                lasteleindex = token[1]
                break
        if(no_of_tables) > 1:
            for token in rh.get_token_atIndexRange(tokens_for_processing, lasteleindex+1, tokens_for_processing[-1][1]):
                if token[0] == tables[n+1].df[0][0].split(" ")[0].replace("：", ""):
                    next_start_index = token[1]
                    no_of_tables = no_of_tables-1
                    # print("no_of_tables", no_of_tables)
                    break
            #print(lasteleindex, next_start_index, "values")
            footertokens = rh.get_between_tokens(tokens_for_processing, lasteleindex+1, next_start_index)
            # print(footertokens)
            footerdata = " ".join(footertokens)
            #print("footer data"+"\n", footerdata)
            lasteleindex = next_start_index
        else:
            footertokens = rh.get_between_tokens(tokens_for_processing, lasteleindex+1, tokens_for_processing[-1][1])
            # print(footertokens)
            footerdata = " ".join(footertokens)
            #print("footer data"+"\n", footerdata)
        string = string+'\n'+footerdata
    return string


if __name__ == "__main__":
    tokens=rh.get_tokens(r"C:\Users\1358527\Documents\rhparser_TI_latest (1)\rhparserv2\inp\3NOD PO (3).pdf")
    print(camelot_extraction(r"C:\Users\1358527\Documents\rhparser_TI_latest (1)\rhparserv2\inp\3NOD PO (3).pdf",tokens))