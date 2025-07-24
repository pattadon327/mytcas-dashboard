# mytcas-dashboard

6510110311 พัทธดนย์ หนุดทอง

## Dashboard for comparing tuition fees for Computer Engineering courses from MyTCAS

แสดงรายละเอียดค่าใช้จ่ายของวิศวกรรมคอมพิวเตอร์และวิศวกรรมปัญญาประดิษฐ์ทั้งหมด โดยใช้ Data ที่ Scraping จาก MyTCAS โดยสามารถ filter ภูมิภาค, มหาวิทยาลัย, หลักสูตร รวมถึงช่วงของค่าใช้จ่ายตามที่ต้องการ

## Coding and Data Details

- tcas1.ipynb : Scraping Data, Cleaning Data
- coe_and_aie_with_major.xlsx : Data จากการ Scraping และผ่านการจัดการแยกคอลัมน์
- mytcas_dash.py : Main Dashboard โดยใช้ path: coe_and_aie_with_major.xlsx
