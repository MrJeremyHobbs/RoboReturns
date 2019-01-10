import requests
import xmltodict
from urllib.parse import urlparse, quote

# classes #####################################################################    
class item_record:
    def __init__(self, barcode, apikey):
        self.item_record = item_record
        
        # generate request url, escape special characters
        url = f"https://api-na.hosted.exlibrisgroup.com/almaws/v1/items?view=label&item_barcode={barcode}&apikey={apikey}"
        encoded_url = quote(url, safe='/:?=&', encoding=None, errors=None)
        self.r = requests.get(encoded_url)
        
        # parse
        self.xml = self.r.text
        self.dict = xmltodict.parse(self.xml)
        
        if self.r.status_code != 200:
            self.found = False
            self.error_msg = self.dict['web_service_result']['errorList']['error'].get('errorMessage')
        else:
            self.found = True
        
            # bib_data
            self.mms_id                       = self.dict['item']['bib_data'].get('mms_id', '')
            self.title                        = self.dict['item']['bib_data'].get('title', '')
            self.author                       = self.dict['item']['bib_data'].get('author', '')
            self.author_short                 = 'not yet implemented'
            self.network_numbers              = self.dict['item']['bib_data'].get('network_numbers', '')
            self.place_of_publication         = self.dict['item']['bib_data'].get('place_of_publication', '')
            self.publisher_const              = self.dict['item']['bib_data'].get('publisher_const', '')
            
            # holding_data
            self.holding_id                   = self.dict['item']['holding_data'].get('holding_id', '')
            self.call_number_type             = self.dict['item']['holding_data'].get('call_number_type', '')
            self.call_number                  = self.dict['item']['holding_data'].get('call_number', '')
            self.accession_number             = self.dict['item']['holding_data'].get('accession_number', '')
            self.copy_id                      = self.dict['item']['holding_data'].get('copy_id', '')
            self.in_temp_location             = self.dict['item']['holding_data'].get('in_temp_location', '')
            self.temp_library                 = self.dict['item']['holding_data'].get('temp_library', '')
            self.temp_location                = self.dict['item']['holding_data'].get('temp_location', '')
            self.temp_call_number_type        = self.dict['item']['holding_data'].get('temp_call_number_type', '')
            self.temp_call_number             = self.dict['item']['holding_data'].get('temp_call_number', '')
            self.temp_policy                  = self.dict['item']['holding_data'].get('temp_policy', '')
            
            # item_data
            self.pid                          = self.dict['item']['item_data'].get('pid', '')
            self.barcode                      = self.dict['item']['item_data'].get('barcode', '')
            self.creation_date                = self.dict['item']['item_data'].get('creation_date', '')
            self.modification_date            = self.dict['item']['item_data'].get('modification_date', '')
            self.base_status                  = self.dict['item']['item_data'].get('base_status', '')
            self.physical_material_type       = self.dict['item']['item_data'].get('physical_material_type', '')
            self.policy                       = self.dict['item']['item_data'].get('policy', '')
            self.provenance                   = self.dict['item']['item_data'].get('provenance', '')
            self.po_line                      = self.dict['item']['item_data'].get('po_line', '')
            self.is_magnetic                  = self.dict['item']['item_data'].get('is_magnetic', '')
            self.arrival_date                 = self.dict['item']['item_data'].get('arrival_date', '')
            self.year_of_issue                = self.dict['item']['item_data'].get('year_of_issue', '')
            self.enumeration_a                = self.dict['item']['item_data'].get('enumeration_a', '')
            self.chronology_i                 = self.dict['item']['item_data'].get('chronology_i', '')
            self.description                  = self.dict['item']['item_data'].get('description', '_blank')
            self.receiving_operator           = self.dict['item']['item_data'].get('receiving_operator', '')
            self.process_type                 = self.dict['item']['item_data'].get('process_type', '')
            self.library                      = self.dict['item']['item_data'].get('library', '')
            
            self.location_dict                = self.dict['item']['item_data'].get('location', '')
            self.location_long                = self.location_dict.get('@desc', '')
            self.location_short               = self.location_dict.get('#text', '')
            
            self.alternative_call_number      = self.dict['item']['item_data'].get('alternative_call_number', '')
            self.alternative_call_number_type = self.dict['item']['item_data'].get('alternative_call_number_type', '')
            self.storage_location_id          = self.dict['item']['item_data'].get('storage_location_id', '')
            self.pages                        = self.dict['item']['item_data'].get('pages', '')
            self.pieces                       = self.dict['item']['item_data'].get('pieces', '')
            self.public_note                  = self.dict['item']['item_data'].get('public_note', '')
            self.fulfillment_note             = self.dict['item']['item_data'].get('fulfillment_note', '')
            
            self.internal_note_1              = self.dict['item']['item_data'].get('internal_note_1', '')
            self.internal_note_2              = self.dict['item']['item_data'].get('internal_note_2', '')
            self.internal_note_3              = self.dict['item']['item_data'].get('internal_note_3', '')
            
            self.statistics_note_1            = self.dict['item']['item_data'].get('statistics_note_1', '')
            self.statistics_note_2            = self.dict['item']['item_data'].get('statistics_note_2', '')
            self.statistics_note_3            = self.dict['item']['item_data'].get('statistics_note_3', '')
            
            self.requested                    = self.dict['item']['item_data'].get('requested', '')
            self.enumeration_a                = self.dict['item']['item_data'].get('enumeration_a', '')
            self.enumeration_b                = self.dict['item']['item_data'].get('enumeration_b', '')
            self.enumeration_c                = self.dict['item']['item_data'].get('enumeration_c', '')
            self.enumeration_d                = self.dict['item']['item_data'].get('enumeration_d', '')
            self.enumeration_e                = self.dict['item']['item_data'].get('enumeration_e', '')
            self.enumeration_f                = self.dict['item']['item_data'].get('enumeration_f', '')
            self.enumeration_g                = self.dict['item']['item_data'].get('enumeration_g', '')
            self.enumeration_h                = self.dict['item']['item_data'].get('enumeration_h', '')
            self.enumeration_i                = self.dict['item']['item_data'].get('enumeration_i', '')
            self.enumeration_j                = self.dict['item']['item_data'].get('enumeration_j', '')
            self.enumeration_k                = self.dict['item']['item_data'].get('enumeration_k', '')
            self.enumeration_l                = self.dict['item']['item_data'].get('enumeration_l', '')
            self.enumeration_m                = self.dict['item']['item_data'].get('enumeration_m', '')
            
            #self.parsed_call_number           = self.dict['item']['item_data']['parsed_call_number'].get('call_no', '')
            
            if self.description == None:
                self.description = ""
                
                
                
class ret:
    def __init__(self):
        self.ret = ret
    
    def post(self, apikey, library, circ_desk, register_in_house_use, mms_id, holding_id, item_pid, xml):
        headers = {
            'Content-Type': 'application/xml', 
            'Charset':'UTF-8', 
            'Authorization': 'apikey {}'.format(apikey)
        }
                   
        params = {
            'op': 'scan',
            'library': library,
            'circ_desk': circ_desk,
            'register_in_house_use': register_in_house_use,
        }
        
        base_url = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1'
        return_url = f"{base_url}/bibs/{mms_id}/holdings/{holding_id}/items/{item_pid}"
        
        r = requests.post(return_url, data=xml.encode('utf-8'), headers=headers, params=params)
        xml = r.text
        dict = xmltodict.parse(xml)
        
        # normalize additional info
        self.additional_info = dict['item']['additional_info']
        self.additional_info = self.additional_info.replace("Item's destination is:", "Destination:")
        
        # successful?
        if r.status_code != 200:
            self.successful = False
            self.error_msg = xml
            self.error_msg = dict['web_service_result']['errorList']['error'].get('errorMessage')
        else:
            self.successful = True