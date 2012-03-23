'''
Script to update the AFH/AFHk Achievements spreadsheet
it depends on pykol and they python bindings to gdata (the google docs api)
you also need a google documents account and a kol account.
fill_spread.py uses a configfile (defaults to fill_spread.cfg overridden with a argument of -c) that looks like:
----
[google]
user=<username>
passwd=<password>
sheet=<spreadsheet name>

[kol]
user=<username>
passwd=<password>
---

also can optionally take an argument -starts to limit to only those runs after starting date
or -ends to those only ending before date (dates are inclusive)

example execution:
python fill_spread.py -c fill_spread.cfg -ends 2011/6/30

when invocated the script logs into google document and examines the spreadsheet
named in the config file, and grabs all the character names in the first column.
it then logs into KoL and scrapes their ascension history finding their fastest
runs of each type, and updates the spreadsheet with the days/turn count and date
of the run.
'''

import gdata.spreadsheet.service

from kol.Session import Session
from kol.request.AscensionHistoryRequest import AscensionHistoryRequest
from kol.request.SearchPlayerRequest import SearchPlayerRequest

import argparse
import ConfigParser
import datetime


def parsedate(d):
    '''
    parse a string in the YYYY/MM/DD format into a date object returning a 
    datetime.date object
    '''
    y,m,d = d.split('/')
    return(datetime.date(int(y),int(m),int(d)))


def _CellsUpdateAction(spr_client, row,col,inputValue,key,wksht_id):
    '''
    update a spreadsheet cell (row/col) with values
    '''
    entry = spr_client.UpdateCell(row=row, col=col, inputValue=inputValue,
            key=key, wksht_id=wksht_id)
    #if isinstance(entry, gdata.spreadsheet.SpreadsheetsCell):
    #    print 'Updated!'


def google_login(username, passwd, sheet_name):
    '''
    login to google docs and select the named spread sheet
    '''
    client = gdata.spreadsheet.service.SpreadsheetsService()
    # Authenticate using your Google Docs email address and password.
    client.ClientLogin(username, passwd)

    feed = client.GetSpreadsheetsFeed()
    for document_entry in feed.entry:
        if document_entry.title.text == sheet_name:
            
            id_parts = document_entry.id.text.split('/')
            curr_key = id_parts[len(id_parts) - 1]
            
            print curr_key
            sfeed = client.GetWorksheetsFeed(curr_key)
            id_parts = sfeed.entry[0].id.text.split('/') # sheet number 1 for the second worksheet
            curr_wksht_id = id_parts[len(id_parts) - 1]
            
    return client, curr_key, curr_wksht_id


def get_names(client, curr_key, curr_wksht_id):
    '''
    look up the names of the ascender's we want records for
    '''
    names = {}
    
    tmp = client.GetCellsFeed(key=curr_key, wksht_id=curr_wksht_id)
    
    for i, entry in enumerate(tmp.entry):
        row = int(entry.cell.row)
        col = int(entry.cell.col)
        if row > 2 and col == 1:
	    name = entry.content.text.replace('\n','').lower()
            #print "(%d,%d) '%s'" % (row, col, name)
            names[name] = row
            
    return names


def update_spread(client, curr_key, curr_wksht_id, row, ascs):
    '''
    iterate over users, and thier fastest runs and update spreadsheet
    '''
    cols = { 'HCNP' : 4,
            'HCB': 6,
            'HCT': 8,
            'HCO': 10,
            'BM': 12,
            'SCNP': 14,
            'SCB': 16,
            'SCT': 18,
            'SCO': 20,
            'HCBHY': 22,
            'SCBHY': 24,
            'HCWSF': 26,
            'SCWSF': 28,
            'SCTrendy': 30,
            'HCTrendy': 32,
            'HCBoris': 34,
            'SCBoris': 36}
    for type in ascs.keys():
        asc = ascs[type]
        if type in cols:
            col = cols[type]

            date = "%d/%d/%d" % (asc['end'].month, asc['end'].day, asc['end'].year)
            _CellsUpdateAction(client, row, col, "%d/%d" %(asc['days'],asc['turns']), curr_key, curr_wksht_id)
            _CellsUpdateAction(client, row, col+1, date, curr_key, curr_wksht_id)
        else: # skip b/c we don't know what column to put it in
            pass
        
def get_uids(s,names):
    '''
    given a list of names, find thier kol user ids
    '''
    uids = {}
    for name in names:
        name = name.lower()
        r = SearchPlayerRequest(s,name)
        responseData = r.doRequest()
        players = responseData['players']
        for player in players:
            if player['userName'].lower() == name:
                uids[name] = player['userId']
    return uids


def get_hist(s,uid, starts=None, ends=None):
    '''
    given a kol user id, look up their fastest runs (within dates)
    '''
    rslt = {}
    ascr = AscensionHistoryRequest(s, uid)
    responseData = ascr.doRequest()
    his = responseData["ascensions"]
    for asc in his:
        if asc['mode'] == 'Casual':
            pass # we don't care about casual
        elif starts is not None and asc['start'] < starts:
            pass # too soon
        elif ends is not None and asc['end'] > ends:
            pass # too late
        else:
            type = lookup_mode_path(asc['mode'],asc['path'])
            if type in rslt:
                if asc['days'] < rslt[type]['days']:
                    rslt[type] = asc
                elif asc['days'] == rslt[type]['days'] and asc['turns'] < rslt[type]['turns']:
                    rslt[type] = asc
            else:
                rslt[type] = asc
    return rslt


def lookup_mode_path(mode,path):
    '''
    lookup table like function for a given mode (HC/SC/BM) and path
    returns the more commonly used HCNP/SCO/HCBHY etc
    '''
    if mode == 'Casual':
        rslt = 'Casual'
    elif mode == 'Softcore':
        rslt = 'SC'
    elif mode == 'Hardcore':
        rslt = 'HC'
    else:
        rslt = 'BM'
        
    if path == 'None':
        if rslt != 'BM':
            rslt = rslt + "NP"

    elif path == "Oxygenarian":
        rslt = rslt + "O"
    elif path == "Boozetafarian":
        rslt = rslt + "B"
    elif path == "Teetotaler":
        rslt = rslt +"T"
    elif path == "Way of the Surprising Fist":
        rslt = rslt +"WSF"
    elif path == "Bees Hate You":
        rslt = rslt+"BHY"
    elif path == "Trendy":
        rslt = rslt + "Trendy"
    elif path == "Avatar of Boris":
        rslt = rslt + "Boris"
    elif path == "Bad Moon":
        pass
    else:
        raise Exception('unknown path: %s' % path)
        
    return rslt

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Update the AFH spreadsheet')
    parser.add_argument('-c', metavar='config_file', type=argparse.FileType('r'),
                nargs='?', default='fill_spread.cfg',
                help='config file for fill_spread.py')
    
    parser.add_argument('-ends', metavar='YYYY/M/D', type=str,
                        nargs='?', default=None,
                        help='date at which runs must finished by')
    parser.add_argument('-starts', metavar='YYYY/M/D', type=str,
                        nargs='?', default=None,
                        help='date at which runs must have started by')
    

    args = parser.parse_args()
    
    if args.ends is not None:
        ends = parsedate(args.ends)
    else:
        ends = None
        
    if args.starts is not None:
        starts = parsedate(args.starts)
    else:
        starts=None

    config = ConfigParser.RawConfigParser(allow_no_value=True)
    config.readfp(args.c)
    
    guser = config.get('google','user')
    gpwd = config.get('google','passwd')
    gsheet = config.get('google','sheet')

    client, curr_key, curr_wksht_id = google_login(guser,gpwd,gsheet)
    
    nrl = get_names(client, curr_key, curr_wksht_id) #get names and thier row
    names = nrl.keys() # list of names
    
    fastest = {}
    # Login to the KoL servers.
    s = Session()
    s.login(config.get('kol','user'), config.get('kol','passwd'))
    
    
    uids = get_uids(s,names) # lookup uids from names
    
    #for each users grab their fastest ascensions and update the spreadsheet with them
    #for name, uid in uids.items():
    pnames = uids.keys()
    pnames.sort()
    for name in pnames:
	uid = uids[name]
        fastest[name] = get_hist(s, uid, starts, ends)
        print name, fastest[name].keys()
        update_spread(client, curr_key, curr_wksht_id, nrl[name],fastest[name])
    s.logout()
