# Static Field Mapping (A:BC)

This project fills the following fields (in order) for each inserted output row:

1. Property Name             -> Start Page!C4
2. Street Address            -> Start Page!C5
3. Country - Region          -> Start Page!C8
4. City/Town                 -> Start Page!C6
5. Zip/Post Code             -> Start Page!C7
6. PropertyType              -> Filter 1
7. Location                  -> Filter 2
8. Food Bev Operator         -> Filter 3
9. Operator                  -> Filter 4
10. Rooms                    -> Start Page!C11
11. Chain/ChainID            -> Filter 5
12. Classification           -> Start Page!F10
13. Management Company       -> Filter 6
14. Owner Company            -> Filter 7
15. YearOpened               -> Filter 8
16. MeetingSpace (SQM)       -> Filter 9
17. MeetingRooms             -> Start Page!F15
18. MeetingMaxCapacity (...) -> Start Page!F16
19. Ski                      -> Filter 10 (Yes/No)
20. Spa                      -> Filter 11 (Yes/No)
21. HealthClub               -> Filter 12 (Yes/No)
22. Golf                     -> Filter 13 (Yes/No)
23. Boutique                 -> Filter 14 (Yes/No)
24. FoodOutlets              -> Filter 15 (text)
25. BeverageOutlets          -> Filter 16 (text)
26. pcd2                     -> Start Page!C7
27. lat                      -> Automation 1 (geocode)
28. long                     -> Automation 2 (geocode)
29. Currency                 -> Start Page!C21

Dynamic transposed data goes into BD:LZ from Projections rows 6:288 for the qualifying column.
