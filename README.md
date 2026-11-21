# Intangible Economy Profiles - Data Visualization 

## Summary
The WIIH Intangible Economy Profiles are the organization's first data visualization product built entirely in-house with AI assistance. They cover 29 economies and were produced without outsourcing any work. Each profile turns harmonized data on intangible investment into a short visual story for policymakers: how much an economy invests in intangible assets, which asset types lead, and how it compares with other economies over time.
The method combines three tools. Reproducible code in R produces the statistics. Excel is used to check every figure. Gen-AI methods were used to produce a master HTML infrastructure that any economy's numbers can be fed into. The first profiles are published as PDFs. An interactive version, where users can scroll, zoom and download the data, comes out in mid-December.

## Background
International organizations usually produce country profiles in one of two ways. They reproduce charts directly from specialized software using automated structured text, or they outsource the editorial publication layout to design agencies. Both routes are costly. Licences and contracts consume budget, and each new edition brings another round of briefs, drafts and revisions that can take months. The profile's structure is often rebuilt every time, even when only the numbers change.
For the WIIH Intangible Economy Profiles, I designed a method that separates these two jobs. Building the structure of each chart, card and time series is demanding, but it only has to be done once. Filling that structure with an economy's data is mechanical. Once creating the data structure, adding an economy or updating an edition takes little more than refreshing a spreadsheet.

## Data sources
All figures come from the Luiss Business School Global INTAN Database, which provides aggregated measures of intangible investment. I collected and harmonized the data for all 29 economies it covers, so that every economy, asset type and year can be compared on the same basis.
The database rests on measuring intangibles through investments. It estimates the value of investment in assets that have no physical form but generate returns over several years, such as software, research, designs, brands and skills. It follows the Corrado–Hulten–Sichel (CHS) framework proxying the total factor of production of employees based on salaries and time spent on an intangible project, categorizing assets into three groups: computerized information, innovative property and economic competencies.

A key feature is the split between measured and unmeasured intangibles. Measured assets are already counted as investment in the System of National Accounts. Unmeasured assets are still recorded as current spending, so their contribution to growth and productivity does not show up in official GDP. Statistical agencies have moved this line over time. The US Bureau of Economic Analysis began treating software as investment in 1999, and in 2013 added R&D and entertainment, literary and artistic originals, which is very much related to the creative industries' copyright and royalties work.

## Intended audience
The profiles are written for WIPO's member states including ministries, national statistical offices, innovation agencies, industry leaders, and academia. They need to see where their economy stands, how it compares with others and which assets drive the trend. Most of these readers are not economists or statisticians, so the profiles are designed for a non-technical audience. Each opens with a few headline numbers, then leads the reader through cards that each make one point: the size of intangible investment, its growth, its composition and its position against other economies. Technical readers, such as researchers and journalists, can go further in the interactive version and download the underlying data.

## Methodology
The work followed six steps. Each step produces an output the next one builds on, and the whole sequence can be repeated for a new edition.
Harmonization and statistics in R. I harmonized the Global INTAN data in RStudio so that all economies and asset types share the same units, years and definitions. Around ten automated code chunks compute the core statistics (year-on-year growth, compound annual growth rates (CAGR), shares by asset type and benchmarks against peers) and draft the charts.
- Validation in Excel. The R output is exported as one spreadsheet per economy. In Excel I cross-checked every figure to the decimal point against the source and the R results, so no number reaches a chart without a second check.
- A master HTML data infrastructure built with AI methods. Using structured prompts, I built one master HTML template that holds every element a profile needs. The chart logic is defined once, from creating axis scales and labels to the layout of each time series and comparison chart. Each economy's data has clearly marked slots in the template, so any AI tool can fill it mechanically without changing the design.
- In parallel, I wrote the cards and key insights for each economy. The goal is a story a policymaker can read in a few minutes: one message per card, plain language, and each number placed beside the chart that shows it.
- Pilot and security checks. The template went through several pilot runs and security reviews with the organization and the design team. These checked accuracy, visual consistency and whether the code was safe to publish. The final code was organized in JavaScript, which drives the interactive features.
- Finally, I filled the master code one spreadsheet at a time, producing a complete profile for each of the 29 economies.
This order puts human judgment where it matters most: the data, the checks and the story. The repetitive part, placing numbers into charts, is automated.

## Outputs
The first versions of the profiles are published as PDFs, suited to the printed report and formal distribution. The interactive version comes out in mid-December. In it, users can:
- scroll through each economy's story card by card;
- zoom into charts and time series;
- download the underlying data;
- take clean screenshots of any chart for presentations and briefings.

Both versions come from the same master infrastructure, so the PDF and the interactive profile show the same numbers.

## Relevance to the creative economy
Intangible investment and the creative economy overlap heavily. Entertainment, literary and artistic originals (films, music, books, broadcasts) are a measured intangible asset in their own right. Design and brand investment, still unmeasured, are produced largely by creative services such as design studios, advertising and market research. Software and databases sit close to the digital side of the creative industries. Much of this investment is made by, or bought from, the cultural industries in each jurisdiction.

This makes the Global INTAN data a practical bridge from national accounts into the creative economy. Creative-economy statistics today are often built from employment counts, trade flows or surveys of copyright-based industries, such as studies following WIPO's own guide on the economic contribution of copyright-based industries. Intangible investment adds a different angle. It treats creative output as capital that keeps producing value over time, measured in monetary terms consistent with GDP.
The method itself transfers directly. The same master infrastructure could produce creative economy profiles by economy or by sector: investment in artistic originals, design and brands, its growth, and the share still unmeasured in official statistics. Because the design is set once, these profiles could serve directors' posts, public editorial reviews, events and social media, with designs strong enough to engage a wide audience.

## Impact
The method shows that data visualization at the standard expected of an organization of WIPO's calibre can be produced in-house at almost no cost. Its impact falls in six areas.
No outsourcing budget was spent. The tools are free or already available in the organization, and the design work was done in-house.
Months of briefs, drafts and agency revisions were replaced by a pipeline where the hard work is done once. After that, a new economy or edition is mainly a matter of populating the data and running the checks.
Every number passes through reproducible code and a cell-by-cell check in Excel before reaching a chart. This removes the transcription errors common in manual layout.
The same structure can take more economies, new indicators or future WIIH editions without being rebuilt, and can be adapted to other WIPO publications.
The organization owns the method, the code and the design. It no longer depends on an outside supplier, and the skills stay in the team.
Interactive, downloadable profiles make the data easier to use and share, raising the visibility of WIPO's work on intangibles among member states, industry and the public.

## Conclusion
This workflow can go well beyond producing profiles. It is a first look at what a Creative Economy Data Model based on GDP and composite measures could become. With alternative data gathered through web scraping, the same approach could power an AI-driven digital platform covering streaming services and large media groups such as YouTube, Spotify and Disney. Working with the design team, I also see sectoral briefs for each economy, with users being able to build their own profiles to suit their needs. This shifts our role from publishing an official number or headline to offering a full infrastructure that academia and policymakers can explore through the CEDM, using official creative economy data sources.
