# Armenian Women's Writing Analysis   

## High-Level Overview of Data

![Sankey chart for data flow](assets/Data_Flowchart.png)

[Back to Top](#armenian-womens-writing-analysis)

## What did English-language Armenian woman newspaper writers write about in the early 21st century?

As a monthly contributor for a bilingual Armenian newspaper in Toronto who covered Dr. Victoria Rowe's PhD dissertation on Armenian woman writers in the late 19th and early 20th centuries (published at the University of Toronto in the year 2000 and available [here](https://utoronto.scholaris.ca/items/03f3f131-4d35-4027-a2fa-ec87ad406335)), I was curious about the patterns in Armenian women's writing in the 21st century.

Since I am an English-language Armenian woman newspaper writer, I decided to narrow the scope of my analysis to my counterparts primarily based in Canada, the US, and Armenia.

This project consisted of several stages:
- scraping data from a representative sample of popular Armenian news sites (read about [here](#data-scraping))
- data cleaning and pre-processing to ensure standardized, uniformly formatted outputs (read about [here](#data-cleaning))
- identifying (likely) women writers (read about [here](#determining-woman-authors-using-ai))
- data visualizations (read about [here](#data-visualization))

The dashboard visualizing the data can be accessed here (link to come). Feel free to observe patterns yourself. I have committed to writing a Torontohye article with my own conclusions, publish date TBD.

[Back to Top](#armenian-womens-writing-analysis)

## Data Scraping

I decided to scrape 3 popular news sites fully and 1 Canadian news site partially. Data scraping was done in December 2024.

The Yerevan-based [EVN Report](https://evnreport.com/about-us/) is a fairly new non-profit, independent online magazine that features long-form deep dive essays on contemporary Armenian issues, written by both local Armenian writers and diasporans.

[The Armenian Weekly](https://armenianweekly.com/) and [The Armenian Mirror-Spectator](https://mirrorspectator.com/) are both older US-based Armenian diasporan papers. The former is traditionally affiliated with the Armenian Revolutionary Federation (ARF) political tendency, while the latter is associated with the Armenian Democratic Liberal Party (ADL), also commonly known as the Ramgavar Party.

In an attempt to create balance between political affiliations (since the Mirror-Spectator had a more thorough public online archive than the Armenian Weekly) and broaden some of the geographic diversity, I also scraped the opinion section of [Horizon Weekly](https://horizonweekly.ca/), a trilingual Canada-based newspaper that is the official publication of Canada's ARF Central Committee. The opinion section featured writing by individual contributors, which was what was needed for my purposes.

Webscrapers were created for each website using the following Python libraries: BeautifulSoup (an HTML parser), Newspaper3k (a library specifically for scraping news articles), requests (to send HTTP requests), and selenium (for navigating websites with unclear pagination patterns).    

Headers and random time delays were employed to simulate human browser requests and prevent rate limiting or being blocked.   

Data was saved to Excel spreadsheets using the openpyxl library. Information saved included each article's title, author, publish date, URL, and keywords generated through natural language processing (NLP).

[Back to Top](#armenian-womens-writing-analysis)

## Data Cleaning

With the article information compiled into Excel spreadsheets, I began to process and prepare the data to be analyzed. To do this, I employed a number of methods: Python scripts using pandas and other libraries, VBA scripts to automate Excel functions, and Excel formulas.

Since I had limited the scope of my analysis to analyzing English-language Armenian women's writing, I had to first remove any articles in French or Armenian. These were removed from the dataset scraped from Horizon Weekly, since the other publications were monolingual English publications.

My end goal was to sort the article data by author gender in order to analyze women's writing. Since it is a person's first name that indicates their likely gender, I needed to isolate the authors' first names.

My prcoess was complicated by the fact that many articles don't have any individual contributors credited. Instead, there is either no author listed; anonymous "guest contributors" credited; the newspaper editorial team collectively credited ("Weekly Staff", "The Armenian-Mirror Spectator"); or an organization credited (ANCA, ARS, Zoryan Institute, etc.). 

Since I wanted to analyze individual women's writing, I had to remove these entries from the dataset. 

Additionally, I needed the standardize the way authors' names were listed. Many authors' names were scraped with honorifics included (Dr., Rev., Fr., etc.). Some authors even had double honorifics listed. I standardized all author names into the same format, with just first names and last names. From this, I was able to isolate the authors' first names and begin sorting by gender.

[Back to Top](#armenian-womens-writing-analysis)

## Determining Woman Authors Using AI 

content to be added

[Back to Top](#armenian-womens-writing-analysis)

## Data Visualization 

content to be added

[Back to Top](#armenian-womens-writing-analysis)

## Credits/Contact

This project was envisioned, researched, and created by Sophia Alexanian (me!). I am an electrical engineering student at the University of Toronto (UofT ECE 2T7+PEY), and a monthly contributor for Torontohye, Toronto's bilingual Armenian newspaper.  
I can be reached here:  
LinkedIn: https://www.linkedin.com/in/sophia-alexanian/   
Email: sophia.alexanian@mail.utoronto.ca  

It was a joy to combine my technical skills with my passion for Armenian women's writing. This project showcases my unique skillset and interests in computer programming, data processing/storytelling, and ethnic women's studies.

Beyond my personal knowledge, since any coding project is made better with external input, I also had some quick check-ins with some other folks I'd like to acknowledge.

Thank you to [Dr. Khatchig Mouradian](https://mesaas.columbia.edu/faculty-directory/khatchig-mouradian/), a professor at Columbia University and the Armenian and Georgian Area Specialist at the Library of Congress, for encouraging work "at the intersection of the Humanities and AI" when we spoke about Victoria Rowe and the future of Armenian studies. 

A BIG thank you to [Khasir Hean](https://www.linkedin.com/in/khasir-hean/), a data scientist/ML engineer (and graduate of UofT's Masters of Applied Computing program), for providing an initial consult about webscrapers. 

Thank you to [Emerson Schryver](https://www.linkedin.com/in/eschry/), a fellow UofT student and CSSU regular, for helping me approach debugging one of the scrapers.

I also extend gratitude to [Robert Zupancic](https://www.linkedin.com/in/robert-zupancic/), fellow UofT engineering student (and my ECE241 lab partner), for doing a final code review and flagging outstanding issues.

P.S. Robert and I created a piano and drum kit simulator for FPGA using Verilog for our final ECE241 project. If you're interested in seeing my work in a hardware description language (as opposed to a high-level programming language), feel free to check out our (censored for academic integrity purposes) work [here](https://github.com/RoZ4/Pianissimo-for-the-FPGA).

[Back to Top](#armenian-womens-writing-analysis)

