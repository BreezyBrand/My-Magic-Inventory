import json
import xlsxwriter

workbook = xlsxwriter.Workbook("Cards.xlsx")
worksheet = workbook.add_worksheet()

all_card_options = ["id"]
debug = False
sample = False
sample_size = 1000
count=0
row=0
skip = 0

class formatted_card:
    def __init__(self,card):        
        self.errors = ""

        try:
            self.id = str(card['id'])
        except:    
            self.errors = self.errors + "Could not set ID. "        
            self.id = ''    
        try:
            self.object = str(card['object'])
        except:
            self.errors = self.errors + "Could not set Object. "
            self.object = ''            
        try:
            self.oracle_id = str(card['oracle_id'])
        except:
            self.errors = self.errors + "Could not set Oracle ID. "
            self.oracle_id = ''
        try:
            self.multiverse_ids = str(card['multiverse_ids'])
        except:
            self.errors = self.errors + "Could not set Multiverse ID. "
            self.multiverse_ids = ''
        try:
            self.mtgo_id = str(card['mtgo_id'])
        except:
            self.errors = self.errors + "Could not set MTGO ID. "
            self.mtgo_id = ''
        try:
            self.mtgo_foil_id = str(card['mtgo_foil_id'])
        except:
            self.errors = self.errors + "Could not set MTGO Foil ID. "
            self.mtgo_foil_id = ''
        try:
            self.tcgplayer_id = str(card['tcgplayer_id'])
        except:
            self.errors = self.errors + "Could not set TCG Player ID. "
            self.tcgplayer_id = ''
        try:
            self.cardmarket_id = str(card['cardmarket_id'])
        except:
            self.errors = self.errors + "Could not set Cardmarket ID. "
            self.cardmarket_id = ''
        try:
            self.name = str(card['name'])
        except:
            self.errors = self.errors + "Could not set Name. "
            self.name = ''
        try:
            self.lang = str(card['lang'])
        except:
            self.lang = ''
        try:
            self.released_at = str(card['released_at'])
        except:
            self.released_at = ''
        try:
            self.uri = str(card['uri'])
        except:
            self.uri = ''
        try:
            self.scryfall_uri = str(card['scryfall_uri'])
        except:
            self.scryfall_uri = ''
        try:
            self.layout = str(card['layout'])
        except:
            self.layout = ''
        try:
            self.highres_image = str(card['highres_image'])
        except:
            self.highres_image = ''
        try:
            self.image_status = str(card['image_status'])
        except:
            self.image_status = ''
        try:
            self.image_uris = str(card['image_uris'])
        except:
            self.image_uris = ''
        try:
            self.mana_cost = str(card['mana_cost'])
        except:
            self.mana_cost = ''
        try:
            self.cmc = str(card['cmc'])
        except:
            self.cmc = ''
        try:
            self.type_line = str(card['type_line'])
        except:
            self.type_line = ''
        try:
            self.oracle_text = str(card['oracle_text'])
        except:
            self.oracle_text = ''
        try:
            self.power = str(card['power'])
        except:
            self.power = ''
        try:
            self.toughness = str(card['toughness'])
        except:
            self.toughness = ''
        try:
            self.colors = str(card['colors'])
        except:
            self.colors = ''
        try:
            self.color_identity = str(card['color_identity'])
        except:
            self.color_identity = ''
        try:
            self.keywords = str(card['keywords'])
        except:
            self.keywords = ''
        try:
            self.legalities = str(card['legalities'])
        except:
            self.legalities = ''
        try:
            self.games = str(card['games'])
        except:
            self.games = ''
        try:
            self.reserved = str(card['reserved'])
        except:
            self.reserved = ''
        try:
            self.foil = str(card['foil'])
        except:
            self.foil = ''
        try:
            self.nonfoil = str(card['nonfoil'])
        except:
            self.nonfoil = ''
        try:
            self.finishes = str(card['finishes'])
        except:
            self.finishes = ''
        try:
            self.oversized = str(card['oversized'])
        except:
            self.oversized = ''
        try:
            self.promo = str(card['promo'])
        except:
            self.promo = ''
        try:
            self.reprint = str(card['reprint'])
        except:
            self.reprint = ''
        try:
            self.variation = str(card['variation'])
        except:
            self.variation = ''
        try:
            self.set_id = str(card['set_id'])
        except:
            self.set_id = ''
        try:
            self.set = str(card['set'])
        except:
            self.set = ''
        try:
            self.set_name = str(card['set_name'])
        except:
            self.set_name = ''
        try:
            self.set_type = str(card['set_type'])
        except:
            self.set_type = ''
        try:
            self.set_uri = str(card['set_uri'])
        except:
            self.set_uri = ''
        try:
            self.set_search_uri = str(card['set_search_uri'])
        except:
            self.set_search_uri = ''
        try:
            self.scryfall_set_uri = str(card['scryfall_set_uri'])
        except:
            self.scryfall_set_uri = ''
        try:
            self.rulings_uri = str(card['rulings_uri'])
        except:
            self.rulings_uri = ''
        try:
            self.prints_search_uri = str(card['prints_search_uri'])
        except:
            self.prints_search_uri = ''
        try:
            self.collector_number = str(card['collector_number'])
        except:
            self.collector_number = ''
        try:
            self.digital = str(card['digital'])
        except:
            self.digital = ''
        try:
            self.rarity = str(card['rarity'])
        except:
            self.rarity = ''
        try:
            self.flavor_text = str(card['flavor_text'])
        except:
            self.flavor_text = ''
        try:
            self.card_back_id = str(card['card_back_id'])
        except:
            self.card_back_id = ''
        try:
            self.artist = str(card['artist'])
        except:
            self.artist = ''
        try:
            self.artist_ids = str(card['artist_ids'])
        except:
            self.artist_ids = ''
        try:
            self.illustration_id = str(card['illustration_id'])
        except:
            self.illustration_id = ''
        try:
            self.border_color = str(card['border_color'])
        except:
            self.border_color = ''
        try:
            self.frame = str(card['frame'])
        except:
            self.frame = ''
        try:
            self.full_art = str(card['full_art'])
        except:
            self.full_art = ''
        try:
            self.textless = str(card['textless'])
        except:
            self.textless = ''
        try:
            self.booster = str(card['booster'])
        except:
            self.booster = ''
        try:
            self.story_spotlight = str(card['story_spotlight'])
        except:
            self.story_spotlight = ''
        try:
            self.edhrec_rank = str(card['edhrec_rank'])
        except:
            self.edhrec_rank = ''
        try:
            self.penny_rank = str(card['penny_rank'])
        except:
            self.penny_rank = ''
        try:
            self.prices = str(card['prices'])
        except:
            self.prices = ''
        try:
            self.related_uris = str(card['related_uris'])
        except:
            self.related_uris = ''
        try:
            self.purchase_uris = str(card['purchase_uris'])
        except:
            self.purchase_uris = ''
        try:
            self.all_parts = str(card['all_parts'])
        except:
            self.all_parts = ''
        try:
            self.promo_types = str(card['promo_types'])
        except:
            self.promo_types = ''
        try:
            self.arena_id = str(card['arena_id'])
        except:
            self.arena_id = ''
        try:
            self.security_stamp = str(card['security_stamp'])
        except:
            self.security_stamp = ''
        try:
            self.card_faces = str(card['card_faces'])
        except:
            self.card_faces = ''
        try:
            self.preview = str(card['preview'])
        except:
            self.preview = ''
        try:
            self.produced_mana = str(card['produced_mana'])
        except:
            self.produced_mana = ''
        try:
            self.watermark = str(card['watermark'])
        except:
            self.watermark = ''
        try:
            self.frame_effects = str(card['frame_effects'])
        except:
            self.frame_effects = ''
        try:
            self.loyalty = str(card['loyalty'])
        except:
            self.loyalty = ''
        try:
            self.printed_name = str(card['printed_name'])
        except:
            self.printed_name = ''        
        try:
            self.game_changer = str(card['game_changer'])
        except:
            self.game_changer = ''        
        try:
            self.printed_type_line = str(card['printed_type_line'])
        except:
            self.printed_type_line = ''                
        try:
            self.printed_text = str(card['printed_text'])
        except:
            self.printed_text = ''        
        try:
            self.color_indicator = str(card['color_indicator'])
        except:
            self.color_indicator = ''        
        try:
            self.tcgplayer_etched_id = str(card['tcgplayer_etched_id'])
        except:
            self.tcgplayer_etched_id = ''        
        try:
            self.content_warning = str(card['content_warning'])
        except:
            self.content_warning = ''        
        try:
            self.flavor_name = str(card['flavor_name'])
        except:
            self.flavor_name = ''        
        try:
            self.attraction_lights = str(card['attraction_lights'])
        except:
            self.attraction_lights = ''        
        try:
            self.variation_of = str(card['variation_of'])
        except:
            self.variation_of = ''        
        try:
            self.life_modifier = str(card['life_modifier'])
        except:
            self.life_modifier = ''        
        try:
            self.hand_modifier = str(card['hand_modifier'])
        except:
            self.hand_modifier = ''        
        try:
            self.defense = str(card['defense'])
        except:
            self.defense = ''

def tryWrite(col,val):    
    if val.startswith("https://"):
        formattedVal = "'" + val[8:]
    elif val.startswith("http://"):
        formattedVal = "'" + val[7:]
    else:
        worksheet.write(row,col,str(val))                
        return
    
    formattedVal = formattedVal[0:65000]
    try:        
        worksheet.write(row,col,str(formattedVal))        
    except:
        print("Could not write value")

def saveCardAttempt(processed_card):
    try:
        keys = card.keys()        
        for i in keys:
            try:            
                keyCol = getKeyCol(i)
                tryWrite(keyCol,card[i])    
            except:
                tryWrite(0,"")
    except:
        print("Could not parse card data")

def getKeyCol(keyId):
    flexCol = all_card_options.index(keyId)
    return flexCol
    
    if keyId=="id": 
        return 0
    if keyId=="object": 
        return 1    
    if keyId=="oracle_id": 
        return 2
    if keyId=="multiverse_ids": 
        return 3
    if keyId=="mtgo_id": 
        return 4
    if keyId=="mtgo_foil_id": 
        return 5
    if keyId=="tcgplayer_id": 
        return 6
    if keyId=="cardmarket_id": 
        return 7
    if keyId=="name": 
        return 8
    if keyId=="lang": 
        return 9
    if keyId=="released_at": 
        return 10
    if keyId=="uri": 
        return 11
    if keyId=="scryfall_uri": 
        return 12
    if keyId=="layout": 
        return 13
    if keyId=="highres_image": 
        return 14
    if keyId=="image_status": 
        return 15
    if keyId=="image_uris": 
        return 16
    if keyId=="mana_cost": 
        return 17
    if keyId=="cmc": 
        return 18
    if keyId=="type_line": 
        return 19
    if keyId=="oracle_text": 
        return 20
    if keyId=="power": 
        return 21
    if keyId=="toughness": 
        return 22
    if keyId=="colors": 
        return 23
    if keyId=="color_identity": 
        return 24
    if keyId=="keywords": 
        return 25
    if keyId=="legalities": 
        return 26
    if keyId=="games": 
        return 27
    if keyId=="reserved": 
        return 28
    if keyId=="foil": 
        return 29
    if keyId=="nonfoil": 
        return 30
    if keyId=="finishes": 
        return 31
    if keyId=="oversized": 
        return 32
    if keyId=="promo": 
        return 33
    if keyId=="reprint": 
        return 34
    if keyId=="variation": 
        return 35
    if keyId=="set_id": 
        return 36
    if keyId=="set": 
        return 37
    if keyId=="set_name": 
        return 38
    if keyId=="set_type": 
        return 39
    if keyId=="set_uri": 
        return 40
    if keyId=="set_search_uri": 
        return 41
    if keyId=="scryfall_set_uri": 
        return 42
    if keyId=="rulings_uri": 
        return 43
    if keyId=="prints_search_uri": 
        return 44
    if keyId=="collector_number": 
        return 45
    if keyId=="digital": 
        return 46
    if keyId=="rarity": 
        return 47
    if keyId=="flavor_text": 
        return 48
    if keyId=="card_back_id": 
        return 49
    if keyId=="artist": 
        return 50
    if keyId=="artist_ids": 
        return 51
    if keyId=="illustration_id": 
        return 52
    if keyId=="border_color": 
        return 53
    if keyId=="frame": 
        return 54
    if keyId=="full_art": 
        return 55
    if keyId=="textless": 
        return 56
    if keyId=="booster": 
        return 57
    if keyId=="story_spotlight": 
        return 58
    if keyId=="edhrec_rank": 
        return 59
    if keyId=="penny_rank": 
        return 60
    if keyId=="prices": 
        return 61
    if keyId=="related_uris": 
        return 62
    if keyId=="purchase_uris": 
        return 63
    if keyId=="all_parts": 
        return 64
    if keyId=="promo_types": 
        return 65
    if keyId=="arena_id": 
        return 66
    if keyId=="security_stamp": 
        return 67
    if keyId=="card_faces": 
        return 68
    if keyId=="preview": 
        return 69
    if keyId=="produced_mana": 
        return 70
    if keyId=="watermark": 
        return 71
    if keyId=="frame_effects": 
        return 72
    if keyId=="loyalty": 
        return 73
    if keyId=="printed_name": 
        return 74

def initializeSheet():
    worksheet.write(0,1,"object")
    worksheet.write(0,0,"id")
    worksheet.write(0,2,"oracle_id")
    worksheet.write(0,3,"multiverse_ids")
    worksheet.write(0,4,"mtgo_id")
    worksheet.write(0,5,"mtgo_foil_id")
    worksheet.write(0,6,"tcgplayer_id")
    worksheet.write(0,7,"cardmarket_id")
    worksheet.write(0,8,"name")
    worksheet.write(0,9,"lang")
    worksheet.write(0,10,"released_at")
    worksheet.write(0,11,"uri")
    worksheet.write(0,12,"scryfall_uri")
    worksheet.write(0,13,"layout")
    worksheet.write(0,14,"highres_image")
    worksheet.write(0,15,"image_status")
    worksheet.write(0,16,"image_uris")
    worksheet.write(0,17,"mana_cost")
    worksheet.write(0,18,"cmc")
    worksheet.write(0,19,"type_line")
    worksheet.write(0,20,"oracle_text")
    worksheet.write(0,21,"power")
    worksheet.write(0,22,"toughness")
    worksheet.write(0,23,"colors")
    worksheet.write(0,24,"color_identity")
    worksheet.write(0,25,"keywords")
    worksheet.write(0,26,"legalities")
    worksheet.write(0,27,"games")
    worksheet.write(0,28,"reserved")
    worksheet.write(0,29,"foil")
    worksheet.write(0,30,"nonfoil")
    worksheet.write(0,31,"finishes")
    worksheet.write(0,32,"oversized")
    worksheet.write(0,33,"promo")
    worksheet.write(0,34,"reprint")
    worksheet.write(0,35,"variation")
    worksheet.write(0,36,"set_id")
    worksheet.write(0,37,"set")
    worksheet.write(0,38,"set_name")
    worksheet.write(0,39,"set_type")
    worksheet.write(0,40,"set_uri")
    worksheet.write(0,41,"set_search_uri")
    worksheet.write(0,42,"scryfall_set_uri")
    worksheet.write(0,43,"rulings_uri")
    worksheet.write(0,44,"prints_search_uri")
    worksheet.write(0,45,"collector_number")
    worksheet.write(0,46,"digital")
    worksheet.write(0,47,"rarity")
    worksheet.write(0,48,"flavor_text")
    worksheet.write(0,49,"card_back_id")
    worksheet.write(0,50,"artist")
    worksheet.write(0,51,"artist_ids")
    worksheet.write(0,52,"illustration_id")
    worksheet.write(0,53,"border_color")
    worksheet.write(0,54,"frame")
    worksheet.write(0,55,"full_art")
    worksheet.write(0,56,"textless")
    worksheet.write(0,57,"booster")
    worksheet.write(0,58,"story_spotlight")
    worksheet.write(0,59,"edhrec_rank")
    worksheet.write(0,60,"penny_rank")
    worksheet.write(0,61,"prices")
    worksheet.write(0,62,"related_uris")
    worksheet.write(0,63,"purchase_uris")
    worksheet.write(0,64,"all_parts")
    worksheet.write(0,65,"promo_types")
    worksheet.write(0,66,"arena_id")
    worksheet.write(0,67,"security_stamp")
    worksheet.write(0,68,"card_faces")
    worksheet.write(0,69,"preview")
    worksheet.write(0,70,"produced_mana")
    worksheet.write(0,71,"watermark")
    worksheet.write(0,72,"frame_effects")
    worksheet.write(0,73,"loyalty")
    worksheet.write(0,74,"printed_name")

def processCard(card):    
    try:        
        processed_card = formatted_card(card)        
        if(debug):
            print(processed_card.errors)
        newKeys = card.keys()
        for i in newKeys:
            if i in all_card_options:
                pass
            else:
                all_card_options.append(i)
    except Exception as e:
        print(f"Error: {e}")

# with open('all-cards.json','r') as file:
#     data = json.load(file)

initializeSheet()
with open('all-cards.json','r',encoding="utf8") as file:    
    line_count = sum(1 for line in file)

with open('all-cards.json','r',encoding="utf8") as file2:        
    for line in file2:        
        if(count%1000 == 0):            
            pct_num = int((count/line_count)*10000)/100
            pct_str = str(pct_num)
            pct =  pct_str+"%"
            print(pct)
        try:
            if(line != "[" and line != "]" and count > skip):                     
                all_but_one = str(line[:-2])
                card = json.loads(all_but_one)
                processed_card = processCard(card)                
                saveCardAttempt(processed_card)
        except Exception as e:            
            print(f"An error ({e}) occured parsing line {str(count)}")            
        count = count + 1
        row = row + 1
        if(sample and count>=sample_size):                        
            workbook.close()
            quit()
    workbook.close()  
    print("100.00%")      
    quit()

# with open('cards.txt','r',encoding="utf8") as file2:    
#     for line in file2:        
#         if(count%10000 == 0):
#             print(count)
#         try:
#             if(line != "[" and line != "]" and count>skip):
#                 all_but_one = str(line[:-2])
#                 card = json.loads(all_but_one)
#                 saveCardAttempt(card)                
#                 row = row+1
#         except:
#             print("An error occured parsing line "+str(count))            
#         count = count + 1            
        # if(count>100):
        #     workbook.close()
        #     quit()


        
        