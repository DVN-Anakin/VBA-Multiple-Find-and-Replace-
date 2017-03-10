# VBA-Multiple-Find-and-Replace
VBA Macro for finding multiple text strings in a given Word document and replacing them with the appropriate strings based on an Excel database.

==================================================================================================
1. USING
==================================================================================================

So, imagine You have a website that You have created from scratch i.e. You are not using any CMS (WP, Prestashop, Drupal, ...) whatsoever. Now, You want to create multiple language versions of your website. First solution that You can think of is obviously to duplicate those .html files and manually translate/rewrite the strings in the source code. This "solution" is surely logical however very laborious indeed - it would take so much of your time, especially if your website has a lot of source files to go through. Another solution is to just copy&paste the source code of each page into Google translator. Of course, a setback is that both Google and Bing translators sucks.

Because I had the same problem, I tried to come up with solme solution. In my case, I was given an Excel database with the original strings and their translations. The database looks like this:

![alt tag](https://github.com/DVN-Anakin/VBA-Multiple-Find-and-Replace-/blob/master/excel-database.png)

So, I have created a VBA macro to deal with this kind of issue. In the given Word document with your source code, It automatically goes through the code and finds the strings located in the left collumn of the database and replace them with their translations from the next collumn.


==================================================================================================
2. BACKLOGS
==================================================================================================

Of course, there are still many backlogs to deal with - for example duplicate strings, very similar strings, etc. I hope that smarter people can use this for their own benefit.
