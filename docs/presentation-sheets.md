# How the Presentation Sheets Work, and How to Use Them

There are five sheets that matter for actually making a decision. Scores is the master ledger, every ticker and every score, the raw material you would rarely act on directly since it feeds the other four. Moonshot Watchlist is Scores filtered down to speculative small bets worth tracking. Insider Feed shows who is actually buying or selling their own company's stock recently, and how much it matters. Earnings Alerts shows what is reporting soon and how it has tended to behave. Buy List is the flagship: quality stocks worth owning long term, and whether now is a good moment price wise.

## Moonshot Watchlist

The script reads Scores, keeps only rows where the tier is Tier4_Moonshot, the ticker is qualified, and Moonshot_Score is a real number, then sorts by Moonshot_Score descending.

An earlier version of this sheet let blank Moonshot_Score rows sit mixed in with real zero to one hundred scores, which was a real bug. A moonshot tier ticker can pass all five disqualifier gates, making it qualified, while still having zero real fundamentals data behind it, meaning it has no verdict and a blank Moonshot_Score. That means there was nothing to judge yet, not that the ticker scored a zero. The fixed version leaves those tickers off the list entirely, so the sheet should now be a clean, ranked list with no blanks.

Moonshot_Score runs from zero to one hundred, computed from revenue growth against peers at fifty percent weight, profit margin against peers at thirty percent, and whether debt relative to assets is improving or worsening at twenty percent. This score is not comparable to Quality_Score or Fundamentals_Score from other tiers, since it is a completely different formula built for pre revenue or early stage companies where cheap relative to peers is close to a meaningless question. Caution_Flags carry over from Scores unchanged.

A higher Moonshot_Score means a more legitimate growth story among the speculative bets on this list, but this is inherently the highest risk tier in the whole system. The idea from the start was small, equal sized positions spread across several of these names, not sizing up just because one number looks impressive. A single moonshot score is not precise enough to bet big on.

## Insider Feed

The script reads the last thirty days of insider transactions and the universe sheet, keeps only open market purchases and sales, since grants, option exercises, and gifts do not carry real investment signal, above a twenty five thousand dollar materiality floor, flags a cluster when two or more different insiders bought or sold the same ticker in the window, and sorts newest first.

Value is the actual dollar size of the transaction. Percent of market cap puts it in context, since a five hundred thousand dollar purchase means something completely different for a fifty million dollar company than a five hundred billion dollar one. The cluster flag distinguishes one person transacting from multiple unrelated insiders doing the same thing around the same time, which is a meaningfully stronger signal in either direction.

A real, sizeable, clustered purchase is one of the more reliable bullish signals available anywhere in this system, since insiders do not buy their own illiquid stock for fun, they buy because they expect it to go up. A sale is much more ambiguous. This data cannot distinguish a routine, pre scheduled sale, tax obligations, diversification, a plan set up months earlier, from a discretionary one made because someone genuinely thinks the stock is about to drop. Do not treat a single sale as automatically bearish. A cluster of unrelated sellers is a real caution flag, while one person selling is often just personal financial planning.

## Earnings Alerts

The script reads the forward looking earnings date already scraped from Finviz, no separate calendar needed, parses text like Jul 28 into a real date by assuming the current year unless that lands more than five days in the past, in which case it assumes next year to handle dates scraped near a year boundary, keeps tickers reporting within the next fourteen days, joins in Volatility, Reliability, and Archetype from Scores, and sorts soonest first.

Days_Until is urgency. Volatility_Score is how big this stock's earnings swings have historically been, magnitude only, in either direction. Reliability_Score is whether it tends to beat or miss. Earnings_Archetype_Code tells you whether the swings sit on top of a good business or a shaky one.

This sheet identifies candidates for an earnings driven swing trade. It does not predict what a specific upcoming report will do. The closest thing to a textbook setup this system currently flags is high volatility, high reliability, and a volatility harvest archetype, reporting soon. A blank volatility or reliability figure means no earnings history exists yet, a recent IPO for example, which is genuinely unknown rather than a signal in either direction, so do not read a blank as safe.

## Buy List

The script reads Scores, keeps rows where the ticker is qualified, the verdict is a strong candidate or a watch, the tier is not Moonshot, and Quality_Score, Value_Score, and Growth_Score are all genuinely real numbers rather than just a passing Fundamentals_Score. It then sorts by Fundamentals_Score descending, so scanning top to bottom actually corresponds to best opportunities first. Finally it writes three live formula columns, current price, percent of 52 week range, and an entry timing label derived from that percentage.

Verdict and Caution_Flags answer two separate questions side by side, not one confusing mixed signal. Verdict answers how strong the fundamentals case is for this business, driven entirely by Fundamentals_Score. Caution_Flags answers whether there is anything else worth knowing about this specific stock, independent of that. They do not move together because they are not measuring the same thing. A watch verdict covers a wide range of Fundamentals_Score, from the mid fifties up to just under seventy, and two different watch stocks can sit at very different points inside that range. That is exactly why the Fundamentals_Score column sits right next to Verdict, so the actual number can tell them apart rather than treating every watch row as identical. Separately, one watch stock might have clean, reliable earnings and no insider red flags, while another watch stock with the same verdict might have unreliable earnings or heavy recent insider selling. Same business quality verdict, different specific risk profile. That is not a contradiction, it is two independent facts about the same stock.

Entry_Timing has nothing to do with verdict. It is purely about where the price sits right now. A price in the bottom quarter of its 52 week range is labeled strong buy undervalued. The middle half is buy now. The top quarter is wait for pullback. Combining both axes is what actually helps decide what to do. A strong candidate sitting at wait for pullback is a genuinely good business you should not chase at today's price. A watch stock sitting at strong buy undervalued might be a decent, not exceptional business trading at a real discount right now. Neither verdict nor entry timing alone tells you what to do. Together, they do.

One honest limitation worth knowing: Entry_Timing only looks at a stock's own 52 week range. It cannot currently tell whether something is cheap because the whole market dipped or cheap because something is specifically wrong with that company. A strong buy undervalued tag is a real, useful signal, but it is not a substitute for knowing why the price is where it is.

## Putting all five sheets together

For a long term buy, start at the Buy List, sorted best first by Fundamentals_Score. Check Entry_Timing for whether now is a good price. Check Caution_Flags for anything specific worth knowing. Then glance at Insider Feed and Earnings Alerts for that same ticker. Recent insider buying adds confidence, and an earnings report in a few days means today's price could move sharply before you have even settled in.

For a speculative small bet, use the Moonshot Watchlist, ranked best first, sized small and spread across several names rather than concentrated in the top one.

For an earnings driven swing trade, use Earnings Alerts, looking for high volatility, high reliability, and a volatility harvest archetype reporting soon. This is a fundamentally different, shorter horizon strategy from everything else here, and it should not be blended with long term Buy List thinking.

Post Earnings Moves is still not built, since it depends on confirming whether the earnings coverage available can reach the full universe or only large cap names.
