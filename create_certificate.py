import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import uuid
import pandas as pd

def main():

    try:
            data = pd.read_csv("database.csv")
    except FileNotFoundError:
        print("File Could Not Found")
        return 1
    
    data.dropna(inplace=True, axis=1)
    data.drop_duplicates(subset=['Emails'], keep='last', inplace=True)
    names = data['Names'].to_list()
    emails = data['Emails'].to_list()

    if os.path.exists('certificates'):
        print("Folder already exists")
    else:
        os.mkdir('certificates')

    df = pd.DataFrame(columns=['Name', 'Mail', 'Uniq-id', 'Created', 'Sended'])

    for (name, email) in zip(names, emails):
        uniq_id = uuid.uuid4()
        df = df._append({'Name': name, 'Mail': email, 'Uniq-id': uniq_id, 'Created': False, 'Sended': False},ignore_index=True)
        try:
            file_path = 'template.jpeg'
            image = Image.open(file_path)
            draw = ImageDraw.Draw(image)

            color = 'rgb(45, 52, 54)'
            font = ImageFont.truetype('ARIAL.TTF', size=80)
            x = 800 - draw.textlength(name,font) / 2
            y = 500;
            draw.text((x, y), name, fill=color, font= font)

            font = ImageFont.truetype('ARIAL.TTF', size=20)
            draw.text((1200, 1100), str(uniq_id), fill=color, font= font)

            cert_dir = 'certificates/'
            cert_path = cert_dir+email+'.pdf'
            image.save(cert_path)
            print(str(email) + " Success")
            df.loc[df['Mail']==email, 'Created'] = True
        except:
            print(str(email) + " Failed")
    
    df.to_excel('data.xlsx', index=False)
main()