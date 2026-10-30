"""
===============================================================================
clean7_v7.py: TURN-SET EXCLUSIONS FROM THE NAMESAKE REVIEW, PLUS TWO FLAGS
===============================================================================
Runs after finalise_v7.py. The namesake review is a hand check, made on 4 Oct
2026, of the turn-set records of the corpus collected for this study: records
where the act's name is used in another sense, or a member's name belongs to
someone else. Its decisions are written into this script.
Reads   harrison_v7_CLEAN6.json  (or harrison_v7_CLEAN6_1.json)
        harrison_v7_CLEAN6.xlsx  (only for its Rejected sheet)
Writes  harrison_v7_CLEAN7.json         every record, full text
        harrison_v7_CLEAN7.xlsx         same records, for browsing
        harrison_v7_MEDIA_PAGES.xlsx    audio and video interview pages, for the
                                        video work
CLEAN6 is never changed. No record is deleted, renumbered or re-dated, and no
existing value is altered except "Turn Set Exclusion" for the 975 records
below. The script checks this before it writes anything.

1. EXCLUSIONS FROM THE NAMESAKE REVIEW (4 Oct 2026)
   975 turn-set records get a reason in "Turn Set Exclusion":
       act name used in another sense   632   (Asia the continent, the film
                                               The Breakfast Club, a horse ...)
       namesake: another person         300
       no usable text                    24   (site menus, photo captions,
                                               one video page)
       passing mention only              19
   Each record is identified by its ID and checked against a fingerprint of
   its URL, so the script stops if it is pointed at a different corpus
   version. The turn set goes from 10,291 to 9,316 records.
   Two new columns record where every exclusion came from:
       "Turn Set Exclusion Source"  finalise_v7 rule (28 Sep 2026) or
                                    namesake review (4 Oct 2026)
       "Turn Set Exclusion Note"    the note written during the review, where
                                    there is one, and how the two undecided
                                    rows were settled
   "In Turn Set" (Yes/No) spells out the rule: high date confidence and no
   exclusion. It is a convenience column; the rule itself is unchanged.

2. RELEVANCE (a flag; nothing is excluded by it)
   How often the record names the act or a nomination-era member:
       "Named In Title"      Yes/No: the act or a member in the headline
       "Relevance Mentions"  mentions in the text: the act's name, members'
                             full names and aliases, and their surname (or a
                             distinctive first name) once that member has been
                             named in full, as in "Hucknall said" or a Q&A
                             label such as "HUCKNALL:"
       "First Mention At"    character position of the first mention (-1 if
                             none), to see where in a long page the act appears
       "Relevance"           named in title / named 2+ times / named once /
                             not found
   A record named once is almost always a passing mention: all 10 checked on 4
   Oct were (a beginner's guide that lists the Human League among others, a
   Kanye review that names Soul II Soul's hit). "Named 2+ times" is mixed:
   roundups and chart lists sit there beside real articles, so which parts of
   a record are about the act is decided at turn level, not here. Other-era
   members are left out of the count, because many are famous for other
   things (Madonna in Breakfast Club records).

3. MEDIA PAGE (a flag; nothing is excluded by it)
   "Media Page" marks pages whose interview is audio or video rather than
   text:
       audio/video site        a podcast, video or radio-player address
                               (YouTube, Vimeo, PBS Video, BBC Sounds, iHeart
                               podcasts, Mixcloud, ...)
       short audio/video page  under 1,500 characters, with an interview word
                               and an audio/video word in the headline or
                               opening ("Audio will be available later today",
                               NAMM video oral histories, podcast episodes)
   These records have little or no text for the turn build, but they are
   leads for the video work, so they are listed in harrison_v7_MEDIA_PAGES.xlsx
   (records excluded as namesakes, other senses, index pages or never naming
   the act are left out of that list).

COLAB (one cell)
  from google.colab import drive
  drive.mount('/content/drive')
  %run /content/drive/MyDrive/harrison_v7/clean7_v7.py

Set DRY_RUN = True to run every check and print the summary without writing.
===============================================================================
"""

import hashlib
import json
import os
import re
import sys
from collections import Counter

import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

VERSION = "clean7_v7 (4 Oct 2026)"
DIR = os.environ.get("HARRISON_DIR", "/content/drive/MyDrive/harrison_v7")
IN_JSON_CANDIDATES = ["harrison_v7_CLEAN6_1.json", "harrison_v7_CLEAN6.json"]
IN_XLSX_CANDIDATES = ["harrison_v7_CLEAN6_1.xlsx", "harrison_v7_CLEAN6.xlsx"]
OUT_JSON = os.path.join(DIR, "harrison_v7_CLEAN7.json")
OUT_XLSX = os.path.join(DIR, "harrison_v7_CLEAN7.xlsx")
OUT_MEDIA = os.path.join(DIR, "harrison_v7_MEDIA_PAGES.xlsx")
DRY_RUN = False

EXPECTED_RECORDS = 13993
EXPECTED_TURN_SET_BEFORE = 10291
EXPECTED_TURN_SET_AFTER = 9316
EXCEL_CELL_CAP = 32000
REJECTED_TEXT_CAP = 1500
SHORT_PAGE = 1500

SOURCE_OLD = "finalise_v7 rule (28 Sep 2026)"
SOURCE_NEW = "namesake review (4 Oct 2026)"
REASONS = {
    "A": "act name used in another sense",
    "N": "namesake: another person",
    "P": "passing mention only",
    "U": "no usable text",
}
OLD_REASONS = {"index or listing page", "never names the act or a member"}
NEW_COLUMNS = ["Turn Set Exclusion Source", "Turn Set Exclusion Note", "In Turn Set", "Relevance",
               "Relevance Mentions", "Named In Title", "First Mention At", "Media Page"]

# ═════════════════════════════════════════════════════════════════════════════
# EXCLUSIONS (record ID, reason code, first 10 characters of the MD5 of the
# record's URL, focal act). Built from the namesake review workbook, 4 Oct 2026.
# ═════════════════════════════════════════════════════════════════════════════
EXCLUSIONS_TSV = """
36	P	49357667ed	Dire Straits
62	P	71b771e1dc	Dire Straits
65	P	6f0d3d9ca1	Dire Straits
123	A	4602e5b67c	Dire Straits
125	A	e21fb5322d	Dire Straits
160	P	60205ea5b6	Dire Straits
199	A	ad4de78de3	Dire Straits
215	A	0908bb8201	Dire Straits
248	A	56c616de0b	Dire Straits
256	A	3b7457a606	Dire Straits
295	A	3f7e4d5827	Dire Straits
296	A	7ccd4a76ad	Dire Straits
297	A	4ad28fcc96	Dire Straits
311	A	cd83099e01	Dire Straits
322	A	baa4a15a10	Dire Straits
329	A	5be243738a	Dire Straits
344	A	77a935221c	Dire Straits
346	A	303fea9204	Dire Straits
361	P	243617cdeb	Dire Straits
410	A	214ef15e15	Dire Straits
711	A	dd1db21feb	The Knack
720	A	80d231d3b9	The Knack
763	A	4a4f926317	The Knack
816	P	53b756f097	The Knack
849	A	eaff8abf35	The Knack
859	P	ae9508262f	The Knack
873	A	c6444fd0f0	The Knack
912	P	cb78982536	The Knack
920	A	110154b8b3	The Knack
931	A	10a50491f8	The Knack
967	A	cd541a6c83	The Knack
968	A	86eead6e66	The Knack
975	A	64e27457e1	The Knack
988	A	2ce4cff3b4	The Knack
989	P	3b8c726cc8	The Knack
998	N	e54b803c23	The Knack
1000	A	1374afa604	The Knack
1002	A	36b5ad428e	The Knack
1040	N	075d6f365b	Amy Holland
1043	N	82d6f4f12c	Amy Holland
1044	N	b3a4ba4902	Amy Holland
1048	N	d4bffb3ddc	Amy Holland
1049	N	9b8a848ccb	Amy Holland
1059	N	d5c3ca3ce4	Amy Holland
1066	N	649869e52e	Amy Holland
1067	N	48de2e9d40	Amy Holland
1068	N	4df65762f6	Amy Holland
1109	N	edd9dc0ab7	Christopher Cross
1126	N	765b6083a5	Christopher Cross
1270	U	1cac538bea	Christopher Cross
1274	N	c45a418d0e	Christopher Cross
1579	U	de007d2a1c	The Pretenders
1580	U	80a2ebd6e3	The Pretenders
1590	A	8de97d80fd	The Pretenders
1680	A	5984e4d0c3	The Pretenders
1714	U	c2f04563fc	The Pretenders
1773	A	e7f90157e0	The Pretenders
1860	N	a4053be91f	The Pretenders
1890	N	d46989389a	The Pretenders
1906	A	572a0b4874	The Pretenders
2093	N	8bd211dddc	Adam and the Ants
2110	N	3962de417c	Adam and the Ants
2111	N	9ca2999648	Adam and the Ants
2116	N	eb58fb5156	Adam and the Ants
2125	N	4b18cabf55	Adam and the Ants
2137	N	c3a279ec21	Adam and the Ants
2144	N	edbeb50981	Adam and the Ants
2148	N	5d72882317	Adam and the Ants
2149	N	ee12614762	Adam and the Ants
2175	N	8e9ec37bb5	Adam and the Ants
2176	N	f2bc509ded	Adam and the Ants
2180	N	8b6ee74689	Adam and the Ants
2181	N	4cc478b8bc	Adam and the Ants
2182	N	207874f773	Adam and the Ants
2186	N	9b901e65b5	Adam and the Ants
2188	N	ee4c867e5e	Adam and the Ants
2193	N	329d378019	Adam and the Ants
2195	N	efcd3dea1a	Adam and the Ants
2196	N	372e21daa3	Adam and the Ants
2201	N	7ea4b1aa71	Adam and the Ants
2202	N	568e795229	Adam and the Ants
2203	N	b335f1f34a	Adam and the Ants
2232	N	bb65625459	James Ingram
2260	N	d48cccaf39	James Ingram
2261	N	5c3a628e24	James Ingram
2262	N	4cd5c785bd	James Ingram
2263	N	ef8f109396	James Ingram
2271	N	b5e4d548f2	James Ingram
2272	N	8dec3c52be	James Ingram
2302	N	1e754dfa57	James Ingram
2332	N	c84a1d1a84	James Ingram
2340	N	f9f32f6fbc	James Ingram
2373	N	da1bab802a	James Ingram
2377	N	4c363dda4c	James Ingram
3399	A	808b4f4fde	Asia
3406	A	b756ba560a	Asia
3408	A	503c30945e	Asia
3410	A	9b59703d64	Asia
3411	A	7a5a1e917f	Asia
3412	A	02d7850b98	Asia
3423	A	62e609af4c	Asia
3440	A	3a4d05add8	Asia
3444	A	f0fb69a893	Asia
3457	A	1f09f1f24e	Asia
3463	A	792f46ac0c	Asia
3467	A	31c200d8e6	Asia
3473	A	549a4ecd26	Asia
3477	A	84215a68fa	Asia
3489	A	a8f5a644e5	Asia
3493	A	749df48385	Asia
3497	A	fb16bb4b9d	Asia
3498	A	f5ee2162b8	Asia
3503	A	9a610a2680	Asia
3504	A	c9c364464e	Asia
3511	A	08200829ac	Asia
3513	A	a3cae68a32	Asia
3515	A	6832f2a7ab	Asia
3519	A	dd3b090e67	Asia
3520	A	21380ef238	Asia
3522	A	07bc982a7f	Asia
3524	A	d00cfe8b01	Asia
3527	A	dec0ae836f	Asia
3528	A	25a61f6d45	Asia
3529	A	baa5918bf6	Asia
3530	A	2d28de0524	Asia
3533	A	c739216ac4	Asia
3534	A	6ffa9e1247	Asia
3537	A	f4bdbf2920	Asia
3541	A	113a35abe0	Asia
3544	A	8389de04b9	Asia
3545	A	0bac1df64f	Asia
3553	A	9e2664d256	Asia
3555	A	64e9f29d07	Asia
3556	A	42ffb6fc3f	Asia
3557	A	83791255bf	Asia
3561	A	e67c13f3cd	Asia
3564	A	be85b65be8	Asia
3565	A	b4b607c208	Asia
3567	A	8135cdbb0f	Asia
3568	A	73c274d564	Asia
3571	A	b9e2dac901	Asia
3573	A	e5c6af4d42	Asia
3576	A	968cc0af63	Asia
3577	A	c8db850de0	Asia
3580	A	f957432277	Asia
3582	A	582a8d1b0e	Asia
3583	A	02548c322c	Asia
3585	A	d7ffed64c4	Asia
3586	A	7623735325	Asia
3588	A	e22a3351d7	Asia
3589	A	fe89a4b76b	Asia
3591	A	0b7e44dbe8	Asia
3592	A	36a81d4280	Asia
3593	A	a5fc7063de	Asia
3594	A	1dba9ad301	Asia
3595	A	1b1b727f93	Asia
3599	A	578b6a6b72	Asia
3601	A	db2ffa42bf	Asia
3602	A	9f90a50881	Asia
3603	A	055423ac3e	Asia
3605	A	e15b0cbfc1	Asia
3606	A	df99a8f38c	Asia
3607	A	08ddf83381	Asia
3608	A	b0da6e78d3	Asia
3610	A	209d70d574	Asia
3611	A	36f64195fb	Asia
3614	A	62b8098f52	Asia
3616	A	b59bcca388	Asia
3618	A	b673bfd8fa	Asia
3620	A	19a3861158	Asia
3621	A	ec65d76746	Asia
3625	A	194646c702	Asia
3626	A	76e770cf93	Asia
3627	A	7a473a2c35	Asia
3630	A	40e1b5464c	Asia
3631	A	95e5239f8f	Asia
3632	A	d34e15d42f	Asia
3633	A	34b928f18e	Asia
3635	A	7173b6961a	Asia
3638	A	6f3f12fe9a	Asia
3639	A	63fb9bf466	Asia
3640	A	bdaf3ce437	Asia
3641	A	92fecf14a8	Asia
3644	A	7f70669c4b	Asia
3645	A	25bf8bc388	Asia
3646	A	32a587e8be	Asia
3647	A	02ab121114	Asia
3648	A	601e8ccfbe	Asia
3651	A	df838cfc55	Asia
3652	A	bb00acbc84	Asia
3655	A	c0efc5cc29	Asia
3656	A	ca6d1219c0	Asia
3658	A	f3938527c1	Asia
3659	A	48a7f1f4ca	Asia
3660	A	54973a9b33	Asia
3661	A	d0077670c9	Asia
3662	A	e5c3bb6113	Asia
3663	A	a9ce2be010	Asia
3667	A	0154f7f7bf	Asia
3671	A	3ef3bf9688	Asia
3672	A	5a616364ee	Asia
3673	A	b2f3f5b2bf	Asia
3676	A	98a48c472e	Asia
3678	A	506483f708	Asia
3679	A	5de997fea9	Asia
3680	A	95f0ff9f47	Asia
3682	A	93e6529fb2	Asia
3685	A	8c83e7db5d	Asia
3686	A	d3c7315bc4	Asia
3688	A	a29db87769	Asia
3689	A	78b42fb4f7	Asia
3690	A	16d9dbf4c1	Asia
3693	A	d1b1e9d0a1	Asia
3695	A	208b0d6c08	Asia
3698	A	5f404ca132	Asia
3700	A	53570673cf	Asia
3704	A	1d59c0fd6a	Asia
3705	A	3ed19769a0	Asia
3706	A	36fb482563	Asia
3709	A	ea008e3cb7	Asia
3710	A	9cd2f3e971	Asia
3711	A	de6d57d478	Asia
3712	A	b8fe856e2d	Asia
3713	A	01bd3dadb1	Asia
3714	A	4a2f20de04	Asia
3719	A	afc210259e	Asia
3720	A	d30d4df8d8	Asia
3723	A	6f76ab4915	Asia
3724	A	d289abcb2c	Asia
3725	A	13c610d1e5	Asia
3732	A	c99f880e67	Asia
3734	A	14750dec7c	Asia
3735	A	3352957199	Asia
3736	A	ae0c6eccdc	Asia
3737	A	9f9afe77b5	Asia
3738	A	a3fc9c0104	Asia
3739	A	b4bf9f7832	Asia
3740	A	e7ec237318	Asia
3742	A	e9f41ae605	Asia
3746	A	5579dbc37a	Asia
3748	A	477599e384	Asia
3754	A	b5f1f2fd69	Asia
3758	A	5d79222367	Asia
3759	A	474bc31c16	Asia
3761	A	a63e2568fa	Asia
3762	A	f6fe6a1e12	Asia
3763	A	355bf6830a	Asia
3764	A	b715b49eee	Asia
3765	A	35c81d121f	Asia
3766	A	5a798b86ad	Asia
3767	A	07d2b5d065	Asia
3768	A	94755f41fb	Asia
3770	A	f6696000f3	Asia
3771	A	36d8da66fd	Asia
3772	A	0b32b481d1	Asia
3773	A	cdc409ff82	Asia
3774	A	ec39025d16	Asia
3776	A	7deac68ca2	Asia
3778	A	21f51ff06d	Asia
3780	A	19c055648b	Asia
3781	A	d7f0716c81	Asia
3785	A	f9764d30d8	Asia
3791	A	ae8981800a	Asia
3796	A	479eb0a921	Asia
3798	A	7843848ca6	Asia
3799	A	d6d59b40b6	Asia
3800	A	16935397cf	Asia
3801	A	ba296fef59	Asia
3803	A	33124330f5	Asia
3804	A	40eee1e4a2	Asia
3807	A	7928db7d16	Asia
4017	A	f639b8e7c5	Men at Work
4019	A	644e483a41	Men at Work
4026	A	79f34671a8	Men at Work
4027	A	78d382faa4	Men at Work
4028	A	1ed8bafce5	Men at Work
4029	A	ec7ade099f	Men at Work
4030	A	7f77e1a2e3	Men at Work
4031	A	7b830752b1	Men at Work
4046	A	9f5278cc82	Men at Work
4047	A	bddefa977d	Men at Work
4052	A	98e6089349	Men at Work
4053	A	25df80f70f	Men at Work
4070	A	1ed2cd3c47	Men at Work
4078	A	e3512488e2	Men at Work
4104	U	bb997bd320	Men at Work
4107	A	e17f006792	Men at Work
4117	A	eca51938f6	Men at Work
4121	N	8cd1124aab	Men at Work
4126	A	a5a47c114f	Men at Work
4128	A	88ea3a8ed1	Men at Work
4129	A	2d76375120	Men at Work
4149	A	79d1afe3c5	Men at Work
4150	A	7bacea78f6	Men at Work
4169	N	4b064c0a91	Men at Work
4170	N	9d38a18eb8	Men at Work
4173	A	cb08e214ec	Men at Work
4174	A	d66c736420	Men at Work
4187	U	f4f292d1ae	Men at Work
4191	A	e6f0152ac9	Men at Work
4229	N	a2ab48d760	Men at Work
4230	A	c30ec6f0b9	Men at Work
4237	A	90f148c3ed	Men at Work
4328	A	c77f4dbafb	Stray Cats
4332	A	9392130d2a	Stray Cats
4379	A	dd024d507e	Stray Cats
4387	A	e7b31d4117	Stray Cats
4403	A	d47dc9c45f	Stray Cats
4416	U	1394a1281a	Stray Cats
4424	A	cb1a2b98e6	Stray Cats
4427	A	9eba42aa2a	Stray Cats
4437	A	851f9ca499	Stray Cats
4461	A	a9dac5d699	Stray Cats
4522	A	3e835b2c06	Stray Cats
4544	A	418974d16e	Stray Cats
4557	A	1de3646352	Stray Cats
4560	U	fa1178e2b1	Stray Cats
4577	A	dc5e5b40a1	Stray Cats
4580	A	8c4abc9e46	Stray Cats
5082	A	c9da52dfd6	Big Country
5083	A	63d5f2c3c2	Big Country
5085	A	e7e5c8c457	Big Country
5086	A	9ddadd2788	Big Country
5087	A	e86e0785cf	Big Country
5091	A	2bcbaf6d01	Big Country
5092	A	3cc7320000	Big Country
5125	A	0a15b0c884	Big Country
5128	A	621ed0e4bf	Big Country
5140	A	8c8da4194d	Big Country
5154	N	11ef45e906	Big Country
5155	N	6437fe4961	Big Country
5176	A	b72674254f	Big Country
5179	N	7337c915b1	Big Country
5189	N	3cedcc5629	Big Country
5193	N	46dd3d048b	Big Country
5218	A	9981dbbf61	Big Country
5281	N	e4b1ac1e67	Big Country
5311	N	9ebaa6e7c5	Big Country
5317	A	59cf42761f	Big Country
5319	A	3f22ef0e9c	Big Country
5340	A	1311925f91	Big Country
5341	N	c42d80b72e	Big Country
5395	A	916c8539d0	Big Country
5443	A	0dc07436f4	Culture Club
5452	U	290237d822	Culture Club
5459	A	236af4a94e	Culture Club
5461	A	7f2bcb5915	Culture Club
5462	A	8064329f88	Culture Club
5463	A	9bacaf066d	Culture Club
5464	A	de28a52b20	Culture Club
5467	A	25d39881f7	Culture Club
5475	A	7988a7193a	Culture Club
5488	A	71b1dfefb7	Culture Club
5516	A	9eab757200	Culture Club
5517	A	05a4e5dce8	Culture Club
5518	A	87be092a65	Culture Club
5519	A	3258505e66	Culture Club
5520	A	8b7c74f2da	Culture Club
5521	A	fbf1221c85	Culture Club
5522	A	8effed1210	Culture Club
5524	A	85161fa12d	Culture Club
5525	A	b12987afc3	Culture Club
5543	A	889a709aeb	Culture Club
5553	A	23385d39f1	Culture Club
5555	U	bf61d99a29	Culture Club
5557	N	1d3c0cf90e	Culture Club
5558	N	214c5ecae6	Culture Club
5603	A	430c8ed2d3	Culture Club
5615	P	4a5a37d5b6	Culture Club
5617	P	b699a41fd2	Culture Club
5715	P	3b088853f7	Culture Club
5723	A	8eb4767db5	Culture Club
5727	U	6386bd9ce3	Culture Club
5735	A	776b0ddb14	Culture Club
5739	N	abeb7969dd	Culture Club
5741	P	9afd6c9874	Culture Club
5745	N	3b5171ca63	Culture Club
5760	N	d799d39110	Culture Club
5802	N	78ccdcab9a	Culture Club
5823	N	d525f9f6a9	Culture Club
6134	N	160dae1339	Eurythmics
6239	N	7379c82940	Eurythmics
6407	N	8c0c918639	Men Without Hats
6435	N	a1acf0c3d8	Men Without Hats
6436	N	98a6bb1a75	Men Without Hats
6442	N	82cc4cb69e	Men Without Hats
6443	N	a4bb44a68d	Men Without Hats
6472	N	a4aa80f99d	Men Without Hats
6485	N	ce38cfdca2	Men Without Hats
6508	N	cfc16c7bbd	Men Without Hats
6551	A	1c8f9a6777	Musical Youth
6556	A	eb43c92cc8	Musical Youth
6558	A	6fa025f5ad	Musical Youth
6559	A	93c4a7a91c	Musical Youth
6569	N	736eae81db	Musical Youth
6571	A	39da1de1cf	Musical Youth
6578	A	c4cdf96a45	Musical Youth
6579	A	7ca796d2fe	Musical Youth
6580	A	c396c2170c	Musical Youth
6586	A	d58462962b	Musical Youth
6592	A	51a2f37f91	Musical Youth
6593	A	a9d0666e99	Musical Youth
6594	A	780a79e85c	Musical Youth
6595	A	b0bc9997df	Musical Youth
6606	N	9f996492b7	Musical Youth
6610	A	7029ec2da8	Musical Youth
6612	A	5c021a9079	Musical Youth
6614	A	1a8c8770e4	Musical Youth
6622	A	46f3624803	Musical Youth
6623	A	c201b05550	Musical Youth
6633	A	7352a30ebb	Musical Youth
6636	A	1fe457854e	Musical Youth
6642	A	01887d3909	Musical Youth
6644	N	f924934c43	Musical Youth
6647	N	1b2e0f6158	Musical Youth
6650	A	c6eb81d565	Musical Youth
6660	A	d74f3bb5a2	Musical Youth
6662	N	c74b002565	Musical Youth
6665	A	c3d0144df0	Musical Youth
6670	A	3267a4923c	Musical Youth
6674	N	bea046016e	Musical Youth
6675	N	d1aa510ca2	Musical Youth
6678	A	7d6e46604e	Musical Youth
6680	N	dc14975bff	Musical Youth
6681	A	b2d364528c	Musical Youth
6685	N	c80546a3b3	Musical Youth
6708	A	06956f58dd	Musical Youth
6713	N	cc030d88a8	Musical Youth
6734	A	783ab0149b	Musical Youth
6737	A	f4aa1417fa	Musical Youth
6741	N	de4fd8caa1	Musical Youth
6742	A	9f68e2d424	Musical Youth
6743	A	b820763744	Musical Youth
6760	N	b722747a02	Musical Youth
6761	N	2d2ec39cba	Musical Youth
6762	N	7cfaa715fa	Musical Youth
6763	N	8dc3ea20be	Musical Youth
6764	N	fe32f4238f	Musical Youth
6766	N	ca9d9cab29	Musical Youth
6782	A	0b34ef9d19	Musical Youth
6786	N	05a66ec3cb	Musical Youth
6787	N	c73ae3fadf	Musical Youth
6788	N	f7725e5b3c	Musical Youth
6790	N	a9c3a77682	Musical Youth
6791	N	351a97cdbf	Musical Youth
6793	N	86c944ec9f	Musical Youth
6794	N	7a8ddfe6f9	Musical Youth
6819	N	d95c476808	Corey Hart
6820	N	23aac9de3e	Corey Hart
6821	N	7be7f0c425	Corey Hart
6822	N	ff0b22716f	Corey Hart
6824	N	baf78bc10a	Corey Hart
6825	N	c736c913a5	Corey Hart
6826	N	1f7c78ba60	Corey Hart
6827	N	4f2cbd6227	Corey Hart
6829	N	a7a462e82b	Corey Hart
6830	N	cec4921675	Corey Hart
6832	N	2f66a66fa5	Corey Hart
6833	N	4b74b93105	Corey Hart
6834	N	7dd25b56ce	Corey Hart
6835	N	2b3dcce312	Corey Hart
6836	N	075b55e58e	Corey Hart
6837	N	1f50a1b10e	Corey Hart
6841	N	4843c1dd7e	Corey Hart
6845	N	65ad3b9727	Corey Hart
6847	N	3d6ccf78a6	Corey Hart
6848	N	f68b2e5175	Corey Hart
6849	N	217dad9637	Corey Hart
6850	N	baace28bde	Corey Hart
6851	N	f00ce49fc4	Corey Hart
6852	N	618783e7c0	Corey Hart
6853	N	0c2833286e	Corey Hart
6856	N	892e42459c	Corey Hart
6857	N	299604699b	Corey Hart
6895	N	2eac65040b	Corey Hart
6922	N	db8508e73f	Corey Hart
7413	N	0a075272b0	Frankie Goes to Hollywood
7418	N	8072cbe893	Frankie Goes to Hollywood
7421	N	6d1add1967	Frankie Goes to Hollywood
7428	N	4a64a009be	Frankie Goes to Hollywood
7473	N	80c9ad7584	Frankie Goes to Hollywood
7475	N	f85726bf62	Frankie Goes to Hollywood
7501	N	5031b2bdb6	Frankie Goes to Hollywood
7518	N	0261323f46	Frankie Goes to Hollywood
7529	N	624e4bbdef	Frankie Goes to Hollywood
7531	N	97ed616b58	Frankie Goes to Hollywood
7533	N	dffbec30b3	Frankie Goes to Hollywood
7535	N	2a220c1086	Frankie Goes to Hollywood
7547	N	dfa494ca4c	Frankie Goes to Hollywood
7576	N	2ccae3deb3	Frankie Goes to Hollywood
7578	N	00f467d547	Frankie Goes to Hollywood
7584	N	72cebfbd61	Frankie Goes to Hollywood
7590	N	0bad71a619	Frankie Goes to Hollywood
7591	N	79206aeb6f	Frankie Goes to Hollywood
7598	N	65f3822418	Frankie Goes to Hollywood
7617	N	bbece49b2d	Frankie Goes to Hollywood
7633	N	e35d3ad65d	Frankie Goes to Hollywood
7641	N	40b3532da2	Frankie Goes to Hollywood
7642	N	9648c50613	Frankie Goes to Hollywood
7654	N	66d748cd35	Frankie Goes to Hollywood
7669	N	281f76fc79	Frankie Goes to Hollywood
7679	N	ff0c68354c	Frankie Goes to Hollywood
7681	N	3669900780	Frankie Goes to Hollywood
7686	N	5edb493933	Frankie Goes to Hollywood
7700	N	c376dd8bf6	Frankie Goes to Hollywood
7701	N	e79f39c5eb	Frankie Goes to Hollywood
7702	N	5da0efac38	Frankie Goes to Hollywood
7703	N	37b06f43ad	Frankie Goes to Hollywood
7704	N	61ed854369	Frankie Goes to Hollywood
8188	N	2fff70df75	Freddie Jackson
8239	N	e31ce2e2e7	Freddie Jackson
8240	N	7267638a2d	Freddie Jackson
8249	N	882df7ad86	Freddie Jackson
8296	N	07d879cae2	Freddie Jackson
8695	N	412bca30e4	Katrina and the Waves
8719	N	0bb9514364	Katrina and the Waves
8723	N	b4b84a7894	Katrina and the Waves
8726	N	b2fd3ee321	Katrina and the Waves
8727	N	2aa6e0896b	Katrina and the Waves
8729	N	fce61b1622	Katrina and the Waves
8737	N	670ba0f242	Katrina and the Waves
8742	N	ec1d50e8d0	Katrina and the Waves
8752	N	d5620903f1	Katrina and the Waves
8758	N	43c3267cc6	Katrina and the Waves
8760	N	9a7737509c	Katrina and the Waves
8761	N	bc6755659e	Katrina and the Waves
8762	N	ca1ea9fdac	Katrina and the Waves
8816	U	69d1364e1c	Sade
8817	U	44cccff8d5	Sade
8821	U	cef61125e5	Sade
8826	A	8aa277a6cd	Sade
8828	A	cff496b5cd	Sade
8829	A	59a9470b12	Sade
8837	U	25ba338f33	Sade
8841	A	89ced858ad	Sade
8843	A	73e97f4456	Sade
8852	A	aea1f79749	Sade
8856	N	1b868d3e4b	Sade
8879	N	6a07f33647	Sade
8911	U	af58d1d809	Sade
8939	P	3dbcc17416	Sade
8989	A	e9dd431b04	Sade
9021	A	30689e9e7c	Sade
9027	P	16d4f0fe06	Sade
9044	U	4e909df5ce	Sade
9055	A	ba658d05bb	Sade
9061	N	3a437d92ea	Sade
9093	A	060317f258	Sade
9114	P	8c5225199e	Sade
9123	A	a6fabc0c0e	Sade
9141	N	bcc3878154	Sade
9148	N	f6c0261f34	Sade
9152	N	44562710fd	Sade
9171	U	a5805ec8e5	Sade
9174	A	9af42c8dad	Sade
9211	A	9c51a65645	a-ha
9224	A	1fa5fac958	a-ha
9250	A	af5c69f96c	a-ha
9252	A	ea276d282d	a-ha
9260	A	80e2b0db03	a-ha
9274	A	0ebe94adab	a-ha
9297	A	943646fdb0	a-ha
9298	A	2103cbd688	a-ha
9309	A	2cec3e1a8b	a-ha
9319	A	ea97385e9f	a-ha
9366	A	d331ff30fc	a-ha
9375	A	90c90e562b	a-ha
9376	A	a4bd066c73	a-ha
9380	A	39990422d8	a-ha
9389	A	aaf440da15	a-ha
9402	A	e24200a439	a-ha
9418	A	cc5508c014	a-ha
9426	A	b21b477a0a	a-ha
9439	P	3bbe33aa27	a-ha
9447	A	d4f458a1d3	a-ha
9471	A	065127d8e2	a-ha
9481	A	938e624ba3	a-ha
9488	A	bbfdf12bb9	a-ha
9492	A	1fac3a62ae	a-ha
9562	N	860b2f8710	Bruce Hornsby and the Range
9666	N	54678b813d	Bruce Hornsby and the Range
9678	N	a41fe73f41	Bruce Hornsby and the Range
9680	N	561b633e5d	Bruce Hornsby and the Range
9681	N	ba12c04487	Bruce Hornsby and the Range
9684	N	7fe7b5eb08	Bruce Hornsby and the Range
9692	N	f9419624c7	Bruce Hornsby and the Range
9754	N	535679c618	Glass Tiger
9757	N	843d1906f8	Glass Tiger
9788	N	a89ca8493c	Glass Tiger
9800	N	cfb4a1a71b	Glass Tiger
9804	N	b37ae07a8a	Glass Tiger
9808	N	6d106f6b90	Glass Tiger
9812	N	306d91915e	Glass Tiger
9826	N	8e5c96fb7e	Glass Tiger
9839	N	3e9bdce784	Glass Tiger
9841	N	72389fd970	Glass Tiger
9846	N	c2ab894bb1	Glass Tiger
9851	N	9b70814349	Glass Tiger
9852	N	5ee2f1649e	Glass Tiger
9853	N	296c807719	Glass Tiger
9854	N	a3c4b2785a	Glass Tiger
9858	N	57ff822c44	Glass Tiger
9861	N	b96cd35bf7	Glass Tiger
9862	N	636c519178	Glass Tiger
9863	N	b7dd877afe	Glass Tiger
9864	N	9a1d6c381c	Glass Tiger
9866	N	4aebd8cbc9	Glass Tiger
9867	N	5de63fe8a2	Glass Tiger
9868	N	1ab4f2769f	Glass Tiger
9919	N	1e09964343	Nu Shooz
9942	N	6185c0dc02	Nu Shooz
9970	N	0e30d096f3	Nu Shooz
9978	N	134534da19	Nu Shooz
9979	N	5b6012b854	Nu Shooz
9984	N	0e301b7f1c	Nu Shooz
10002	N	70a18b753b	Nu Shooz
10006	N	4e47a6c92c	Nu Shooz
10007	N	3d216df6c7	Nu Shooz
10012	N	13a3ab9344	Nu Shooz
10015	N	b69e2a440f	Nu Shooz
10016	N	8b07e835d3	Nu Shooz
10018	N	242aeeed02	Nu Shooz
10077	U	ff5874c8ad	Simply Red
10085	A	f296f9f1f5	Simply Red
10088	A	80a2ebd6e3	Simply Red
10094	A	bf2e6d3007	Simply Red
10096	A	24fa01b5f8	Simply Red
10116	U	294cfbe210	Simply Red
10141	A	6198e11d26	Simply Red
10154	A	734e296e13	Simply Red
10169	A	5869eb5967	Simply Red
10190	P	b485070e6e	Simply Red
10257	U	71a5382a3e	Simply Red
10304	P	8c5225199e	Simply Red
10451	A	8dac6feb88	Breakfast Club
10454	A	8307db59fa	Breakfast Club
10455	A	9810d6740c	Breakfast Club
10456	A	c223a2be1e	Breakfast Club
10457	A	f36b310a12	Breakfast Club
10458	A	e6b8f34035	Breakfast Club
10459	A	4938ed10bd	Breakfast Club
10460	A	21af842aec	Breakfast Club
10462	A	c7ad6e1814	Breakfast Club
10463	A	c0205d5762	Breakfast Club
10466	A	6bd7dc4bf4	Breakfast Club
10468	A	da77d95fc7	Breakfast Club
10472	A	30275b3fd2	Breakfast Club
10473	A	f2c5ea6d35	Breakfast Club
10480	A	2398aa11fb	Breakfast Club
10481	A	f0968f05b2	Breakfast Club
10485	A	8f2187d15b	Breakfast Club
10486	A	93007213f2	Breakfast Club
10487	A	3f5091563f	Breakfast Club
10488	A	3316ae8697	Breakfast Club
10490	A	99e1b4bbda	Breakfast Club
10494	A	00290dccf6	Breakfast Club
10495	A	287b1e081c	Breakfast Club
10496	A	60f7afc5aa	Breakfast Club
10497	A	d3da4ff985	Breakfast Club
10498	A	ba5b191d53	Breakfast Club
10499	A	8465b1beb4	Breakfast Club
10500	A	490c7d667d	Breakfast Club
10501	A	6d4c5fed59	Breakfast Club
10502	A	7658959d31	Breakfast Club
10504	A	376558a0b4	Breakfast Club
10505	A	f0a892d728	Breakfast Club
10507	A	2645de075b	Breakfast Club
10509	A	43d0df342f	Breakfast Club
10511	A	94e2a0f8e5	Breakfast Club
10512	A	94d6db752d	Breakfast Club
10515	A	11f7f7814d	Breakfast Club
10516	A	5585f77a9e	Breakfast Club
10517	A	c7fd46bc34	Breakfast Club
10519	A	473352c3f5	Breakfast Club
10523	A	7be1c2cfa9	Breakfast Club
10524	A	5322f344bc	Breakfast Club
10525	A	18c2eb6e88	Breakfast Club
10528	A	4801a69077	Breakfast Club
10532	A	6bb1f29dcb	Breakfast Club
10534	A	378b01b90b	Breakfast Club
10537	A	cbcfadd6f0	Breakfast Club
10538	A	6849f93e84	Breakfast Club
10541	A	58d7f6b37d	Breakfast Club
10542	A	337e1af8b6	Breakfast Club
10544	N	dfe676190a	Breakfast Club
10545	N	be925c87d2	Breakfast Club
10546	N	624029ab48	Breakfast Club
10547	N	7a772f8c06	Breakfast Club
10548	N	fb8dbbd829	Breakfast Club
10549	N	33b1a595cb	Breakfast Club
10550	N	75aec84b86	Breakfast Club
10552	A	ba8f844ebb	Breakfast Club
10555	A	2e53da9c52	Breakfast Club
10556	A	fc31ac40e3	Breakfast Club
10557	A	e76ecdb890	Breakfast Club
10558	N	a2f6f205af	Breakfast Club
10562	A	080ad4c976	Breakfast Club
10563	A	1784e6fb1d	Breakfast Club
10564	A	ea6c5d4ccb	Breakfast Club
10565	A	a2cff2515c	Breakfast Club
10566	A	d7d7fd81fc	Breakfast Club
10567	A	fab86f0bcf	Breakfast Club
10569	A	3b09d75ae2	Breakfast Club
10570	A	e37d9d6a32	Breakfast Club
10571	A	98bf96f5fc	Breakfast Club
10572	A	1daadff195	Breakfast Club
10574	A	f166583ad6	Breakfast Club
10577	A	6ed1afb56c	Breakfast Club
10578	A	addf2fb61a	Breakfast Club
10579	A	0f21062336	Breakfast Club
10581	N	893f0e55bc	Breakfast Club
10583	A	91adf8cfef	Breakfast Club
10584	A	79e376bb40	Breakfast Club
10586	A	bf9ac9b40c	Breakfast Club
10587	A	eefe786d44	Breakfast Club
10588	A	40ebc12dd2	Breakfast Club
10589	A	ef63239c11	Breakfast Club
10590	A	34eaea7b41	Breakfast Club
10591	A	7601821496	Breakfast Club
10592	A	cdd3a85848	Breakfast Club
10593	A	cb0c2c9d6f	Breakfast Club
10594	A	23b2832589	Breakfast Club
10595	A	9ea9f6c1f3	Breakfast Club
10597	A	22c098d918	Breakfast Club
10598	A	a29c170f50	Breakfast Club
10599	A	6abca8071c	Breakfast Club
10601	A	f3db9f5664	Breakfast Club
10602	A	b8c3660578	Breakfast Club
10603	A	1047be3d1b	Breakfast Club
10604	A	29242057ba	Breakfast Club
10605	A	d0730741d2	Breakfast Club
10606	A	db081684b4	Breakfast Club
10607	A	cabfeb004b	Breakfast Club
10608	N	9f7cb971fe	Breakfast Club
10609	N	4f6edb029a	Breakfast Club
10610	A	b06ded2560	Breakfast Club
10611	A	74eb33a503	Breakfast Club
10612	N	ffa4772cd3	Breakfast Club
10614	N	d3be8ed9e0	Breakfast Club
10615	A	16c8aa51c5	Breakfast Club
10616	A	3b2f894716	Breakfast Club
10617	A	2c15e77332	Breakfast Club
10618	A	0cc8c35b14	Breakfast Club
10619	A	d5dfed551b	Breakfast Club
10620	A	5211ec6097	Breakfast Club
10622	A	d219dfad30	Breakfast Club
10623	A	92c5de9646	Breakfast Club
10624	A	08ee1b42ee	Breakfast Club
10625	A	2cd57eaf27	Breakfast Club
10626	A	08683131e7	Breakfast Club
10627	A	3a48e48e5b	Breakfast Club
10628	A	ffd85df6fc	Breakfast Club
10629	A	8e11a9491f	Breakfast Club
10630	A	7db9feb72b	Breakfast Club
10631	A	4716ad25cf	Breakfast Club
10633	N	acfb28fa9a	Breakfast Club
10634	N	0d2a892110	Breakfast Club
10636	A	aabf6e5c94	Breakfast Club
10638	A	52728817dd	Breakfast Club
10640	A	68fb7affb3	Breakfast Club
10641	A	a9f31f390f	Breakfast Club
10642	A	8a4c04279d	Breakfast Club
10643	N	78b0e42ed6	Breakfast Club
10644	A	120623c05e	Breakfast Club
10648	A	4525209f24	Breakfast Club
10649	A	3b47f57bc6	Breakfast Club
10650	A	c7162bc5c2	Breakfast Club
10652	A	f5a60ac3f6	Breakfast Club
10653	A	7ae32053dd	Breakfast Club
10654	A	7b89b4f4a0	Breakfast Club
10655	A	923288041e	Breakfast Club
10660	A	bfcc1e6627	Breakfast Club
10664	A	9095e04eee	Breakfast Club
10666	N	373adba269	Breakfast Club
10667	N	b4f2284402	Breakfast Club
10668	A	b3bf9ee297	Breakfast Club
10672	A	ac15f964a2	Breakfast Club
10673	A	c1a2b84a7c	Breakfast Club
10677	A	542752f55e	Breakfast Club
10678	A	dda39ed451	Breakfast Club
10679	A	4910004c83	Breakfast Club
10682	A	33d8397c1f	Breakfast Club
10686	A	b21169c867	Breakfast Club
10688	A	1112ede7c7	Breakfast Club
10689	A	36294aab15	Breakfast Club
10693	A	eb1a5d2f90	Breakfast Club
10694	A	1dc30eb515	Breakfast Club
10695	A	708089f5df	Breakfast Club
10696	N	b4145cd14a	Breakfast Club
10698	A	22bff993f0	Breakfast Club
10704	A	1aa8f290b3	Breakfast Club
10706	N	d1d7f99760	Breakfast Club
10709	A	183e0d9317	Breakfast Club
10715	A	47683ec6b7	Breakfast Club
10716	A	ba54aaaf1f	Breakfast Club
10718	A	ea9864d559	Breakfast Club
10721	A	be54247e18	Breakfast Club
10723	A	2d30efb7c9	Breakfast Club
10724	A	5a8591d5b4	Breakfast Club
10725	A	d821d51005	Breakfast Club
10727	A	78907cf82b	Breakfast Club
10728	A	d529d2bac2	Breakfast Club
10729	A	00319785ff	Breakfast Club
10730	A	5067548a9c	Breakfast Club
10732	A	4502bdfbf2	Breakfast Club
10734	A	423ab4751e	Breakfast Club
10736	A	f2a0cbe6f5	Breakfast Club
10737	A	5b4b572df7	Breakfast Club
10739	A	8dda468d4e	Breakfast Club
10745	A	a76e909b50	Breakfast Club
10746	A	3def9a36be	Breakfast Club
10747	A	adacb4ec41	Breakfast Club
10748	N	b905f3f320	Breakfast Club
10749	A	bc82295613	Breakfast Club
10751	N	a0d9006a18	Breakfast Club
10753	A	49abf1b048	Breakfast Club
10754	A	a0af5b9664	Breakfast Club
10756	A	25c35bf554	Breakfast Club
10757	A	88d7622f6b	Breakfast Club
10759	A	a6dd53c123	Breakfast Club
10761	A	d1f1d1827a	Breakfast Club
10763	A	3208a35a0a	Breakfast Club
10764	A	898a51433a	Breakfast Club
10765	A	33f7c6d0d6	Breakfast Club
10769	A	bc72c53e3b	Breakfast Club
10770	A	2aa6ca66e5	Breakfast Club
10772	A	d5c7c57322	Breakfast Club
10774	A	dceaf9ca30	Breakfast Club
10777	A	e90a2f42f7	Breakfast Club
10780	A	cca5fb7aee	Breakfast Club
10781	A	9ba01cfc32	Breakfast Club
10782	A	5043d85cbb	Breakfast Club
10783	A	c8b65b9340	Breakfast Club
10784	A	cad09cd845	Breakfast Club
10785	A	dad8ff3006	Breakfast Club
10787	A	aaa0d758a2	Breakfast Club
10789	A	28cf51d847	Breakfast Club
10790	A	b9d2a51f1d	Breakfast Club
10796	A	2aacb5d167	Breakfast Club
10800	A	e5001a6ca5	Breakfast Club
10801	A	996500be37	Breakfast Club
10802	N	846b1c4c72	Breakfast Club
10803	N	d3433bbe6e	Breakfast Club
10805	A	e1bfd255a7	Breakfast Club
10806	A	d9e8fa4a78	Breakfast Club
10808	A	2d45794c67	Breakfast Club
10809	A	a5bed14bf8	Breakfast Club
10810	A	b4bf9f7832	Breakfast Club
10812	A	1ea5cbe522	Breakfast Club
10813	A	ba16cbc0d6	Breakfast Club
10814	N	c36381cf17	Breakfast Club
10815	A	0397c55538	Breakfast Club
10820	A	947188ffc1	Breakfast Club
10822	A	79ed42984e	Breakfast Club
10823	A	a0b887cbd0	Breakfast Club
10825	A	06d62dc4b8	Breakfast Club
10826	A	6d60b12877	Breakfast Club
10827	N	51fcc08552	Breakfast Club
10828	A	4247f16216	Breakfast Club
10829	A	c14681a201	Breakfast Club
10830	A	0de73002bb	Breakfast Club
10831	A	d295db9e7e	Breakfast Club
10835	N	a553d83670	Breakfast Club
10836	A	c2dcb5b769	Breakfast Club
10838	N	fbfe1daa23	Breakfast Club
10839	A	993ba064dd	Breakfast Club
10841	A	743ce1ec16	Breakfast Club
10887	A	a48ca1f701	Cutting Crew
10891	A	c2123e8cdf	Cutting Crew
10896	A	c325a79807	Cutting Crew
10901	A	027aa39de4	Cutting Crew
10904	A	17e799e375	Cutting Crew
10912	N	407c8579c5	Cutting Crew
10915	A	901baeb9c8	Cutting Crew
10952	A	341a02fe6f	Cutting Crew
10956	A	d54588a9f0	Cutting Crew
10957	A	50b0a297cb	Cutting Crew
10963	U	c27f2b9efd	Cutting Crew
10967	A	038a1141eb	Cutting Crew
10976	A	e63721a945	Cutting Crew
11006	A	4d16f88a6e	Cutting Crew
11025	A	64a988eef5	Cutting Crew
11027	N	9af9fb85ea	Cutting Crew
11031	A	b74cab4f3d	Cutting Crew
11037	N	42e16cbdcb	Cutting Crew
11066	U	b0dba84c04	Cutting Crew
11387	N	fa287650f9	Swing Out Sister
11411	N	dde0850627	Swing Out Sister
11418	N	455d80bd76	Swing Out Sister
11421	N	1e778ebe3f	Swing Out Sister
11440	N	dee3a67ec8	Swing Out Sister
12077	A	ab1ebe8beb	Take 6
12078	A	8ca9b0f33a	Take 6
12097	A	1fa852a9ed	Take 6
12099	A	55dcd8e284	Take 6
12110	A	232f395096	Take 6
12147	N	bd7792012b	Take 6
12149	A	5376a39c83	Take 6
12167	A	801e12fa2c	Take 6
12169	A	137780838b	Take 6
12173	N	6682a720d1	Take 6
12179	N	9b62cf845f	Take 6
12185	U	c4fa93574d	Take 6
12186	A	9c98d7e9d6	Take 6
12187	A	3633c19b62	Take 6
12188	N	6b83e52281	Take 6
12189	A	ceb51007a1	Take 6
12196	A	347ecbfea0	Take 6
12200	A	093aed2acd	Take 6
12214	A	31b1e8b3c3	Take 6
12215	A	b6aeb2608d	Take 6
12218	A	038eefd71b	Take 6
12226	A	51970153b4	Take 6
12228	A	5b05001752	Take 6
12231	A	41d325fee0	Take 6
12232	A	8d32dde222	Take 6
12233	A	a96affde12	Take 6
12235	A	abaff02429	Take 6
12236	A	1bf6b77e81	Take 6
12239	A	922182ecf7	Take 6
12240	A	18ab3cfec2	Take 6
12250	A	3820fef820	Take 6
12255	A	33a15b0a5e	Take 6
12260	A	a17744908a	Take 6
12264	A	9d7242a626	Take 6
12267	A	35138fe4ed	Take 6
12268	A	7b378cff45	Take 6
12269	A	b1946de6d0	Take 6
12275	A	ce7035a730	Take 6
12276	A	f9327f699f	Take 6
12277	N	a52c2ebd00	Take 6
12278	A	c896a4bcab	Take 6
12282	N	07df4635a4	Take 6
12284	A	136d2f3470	Take 6
12285	A	de9a36d221	Take 6
12287	A	baf12431eb	Take 6
12288	A	484819d6cd	Take 6
12289	A	3cb7e8aed5	Take 6
12290	A	a5cba14440	Take 6
12299	A	177bf724e3	Take 6
12301	A	b5adaf4c9f	Take 6
12303	A	0a6ad7fd9e	Take 6
12320	A	923d28185f	Take 6
12322	A	2cd180942a	Take 6
12332	N	0064dd7704	Take 6
12333	N	b533414f94	Take 6
12334	N	97540f318a	Take 6
12335	N	216d5090a1	Take 6
12336	N	f638fc49ae	Take 6
12337	N	05ea173673	Take 6
12338	A	d159fe10a2	Take 6
12339	A	b433e17622	Take 6
12343	A	53e4cdb0f5	Take 6
12345	N	f68d0b5e92	Take 6
12362	A	0af360c6ea	Take 6
12404	N	9f0eebfc1f	Toni Childs
12407	N	7561ecb47b	Toni Childs
12413	N	f63b05d863	Toni Childs
12419	N	227b40f53c	Toni Childs
12423	N	402be3d6ec	Toni Childs
12431	N	9c2c8085c2	Toni Childs
12434	N	ce8c962a25	Toni Childs
12446	N	d884a83165	Toni Childs
12614	N	f35968a9b3	Tracy Chapman
12750	N	8bfc2c5514	Vanessa Williams
12766	N	d2d8b30aec	Vanessa Williams
12781	N	431a8df381	Vanessa Williams
12838	N	e1134332cf	Vanessa Williams
12846	N	87522d49d1	Vanessa Williams
12848	N	8f86c4979b	Vanessa Williams
12852	N	17fa6ae6a6	Vanessa Williams
12861	N	1c7aa9897f	Vanessa Williams
12864	N	c381159a46	Vanessa Williams
12897	N	223e2eaeed	Vanessa Williams
12909	N	e198b04f1a	Vanessa Williams
"""

# Notes carried into "Turn Set Exclusion Note"
NOTES = {
    36: "Review note: Dire Straits is mentioned by another musician entirely in passing",
    62: "Review note: Dire Straits is mentioned by another musician entirely in passing",
    65: "Review note: Dire Straits is mentioned by another musician entirely in passing",
    160: "Review note: Dire Straits is mentioned by the author entirely in passing",
    215: "Review note: Dire Straits mentioned in a word game",
    256: "Review note: This article is about something else entirely",
    329: "Review note: This article is about something else entirely",
    361: "Review note: Dire Straits is mentioned by the author in passing",
    711: "Review note: This article is about something else entirely",
    816: "Review note: The Knack mentioned in passing",
    859: "Review note: The Knack mentioned in passing",
    873: "Review note: The article is about something else",
    912: "Review note: The Knack mentioned in passing",
    989: "Review note: The Knack mentioned in passing",
    1270: "Review note: This is a video interview of Christopher Cross; Resolved 4 Oct from the stored text: the page holds only site menus; the interview itself is a video (link added to the media list).",
    1579: "Review note: No relevant content",
    1580: "Review note: No relevant content",
    1590: "Review note: Paywalled, could not be accessed; Stored text checked 4 Oct: an LA Times piece about house-hunting, not the band.",
    1714: "Review note: No relevant content",
    1906: "Review note: This article is about something else entirely",
    4027: "Review note: This article is about something else entirely",
    4028: "Review note: This article is about something else entirely",
    4029: "Review note: This article is about something else entirely",
    4031: "Review note: This article is about something else entirely",
    4053: "Review note: This article is about something else entirely",
    4070: "Review note: This article is about something else entirely",
    4078: "Review note: This article is about something else entirely",
    4104: "Review note: No relevant content",
    4107: "Review note: This article is about something else entirely",
    4117: "Review note: This article is about something else entirely",
    4129: "Review note: This article is about something else entirely",
    4187: "Review note: No relevant content",
    4191: "Review note: This article is about something else entirely",
    4416: "Review note: This page alone is not helpful - but this is the website of the band. Other pages in this website would definitely be helpful",
    4560: "Review note: No usable content",
    5092: "Review note: This article is about something else entirely",
    5125: "Review note: This article is about something else entirely",
    5155: "Review note: This is a different person unrelated to Big Country",
    5176: "Review note: This article is about something else entirely",
    5218: "Review note: This article is about something else entirely",
    5317: "Review note: This article is about something else entirely",
    5340: "Review note: This article is about something else entirely",
    5452: "Review note: No usable content",
    5475: "Review note: This article is about something else entirely",
    5543: "Review note: This article is about something else entirely",
    5553: "Review note: This article is about something else entirely",
    5555: "Review note: No usable content",
    5615: "Review note: Mentioned in passing",
    5617: "Review note: Mentioned in passing",
    5715: "Review note: Mentioned in passing",
    5723: "Review note: This article is about something else entirely",
    5727: "Review note: No usable content",
    5741: "Review note: Mentioned in passing",
    5760: "Review note: This article is about someone else",
    8816: "Review note: No usable content",
    8817: "Review note: No usable content",
    8821: "Review note: No usable content",
    8837: "Review note: No usable content",
    8856: "Review note: This article is about someone else",
    8911: "Review note: No usable content",
    8939: "Review note: Mentioned in passing",
    9027: "Review note: Mentioned in passing",
    9044: "Review note: No usable content",
    9055: "Review note: This article is about something else entirely",
    9114: "Review note: Mentioned in passing",
    9123: "Review note: This article is about something else entirely",
    9171: "Review note: No usable content",
    9174: "Review note: This article is about something else entirely",
    9252: "Review note: This article is about something else entirely",
    9439: "Review note: A-ha mentioned in passing",
    9481: "Review note: This article is about something else entirely",
    9492: "Review note: This article is about something else entirely",
    10077: "Review note: No usable content",
    10085: "Review note: Paywalled; could not be accessed; Stored text checked 4 Oct: a photo gallery about red hair, not the band.",
    10088: "Review note: Unrelated content",
    10094: "Review note: Paywalled; could not be accessed; Stored text checked 4 Oct: a beer called Simi Simply Red, not the band.",
    10096: "Review note: This article is about something else entirely",
    10116: "Review note: No usable content",
    10190: "Review note: Paywalled; could not be accessed; Stored text checked 4 Oct: a record review that mentions 'pre-Simply Red' musicians in passing.",
    10257: "Review note: No usable content",
    10304: "Review note: Simply Red mentioned in passing",
    10814: "Review note: The link does not work, so the article could not be checked; Resolved 4 Oct from the stored text: an obituary for a drummer named Gary Burke; the Breakfast Club is never mentioned.",
    10901: "Review note: Paywalled; could not be accessed; Stored text checked 4 Oct: a racehorse called Cutting Crew.",
    10963: "Review note: These are just images; no text",
    11006: "Review note: This article is about something else",
    11066: "Review note: No usable content",
    12185: "Review note: No usable content",
}

# Other leads for the media list: pages without usable text that still point
# somewhere useful
LEADS = {
    4416: "Official band website: this welcome page has no content, but its other pages may hold "
          "news, interviews or a biography (note from the review).",
}

# ═════════════════════════════════════════════════════════════════════════════
# NAMES FOR THE RELEVANCE COUNT: the act's name, and every nomination-era
# member with the aliases confirmed in the line-up review (4 Oct 2026)
# ═════════════════════════════════════════════════════════════════════════════
ACT_NAME = {  # (pattern, case-sensitive?) as in finalise_v7
    "Dire Straits": (r"dire straits", False), "Rickie Lee Jones": (r"rickie lee jones", False),
    "The Knack": (r"\b(?:[Tt]he )?(?:Knack|KNACK)\b", True), "Amy Holland": (r"amy holland", False),
    "Christopher Cross": (r"christopher cross", False), "Irene Cara": (r"irene cara", False),
    "Robbie Dupree": (r"robbie dupree", False), "The Pretenders": (r"\bpretenders\b", False),
    "Adam and the Ants": (r"adam (?:and|&) the ants|adam ant\b", False),
    "James Ingram": (r"james ingram", False), "Luther Vandross": (r"vandross", False),
    "Sheena Easton": (r"sheena easton", False), "The Go-Go's": (r"go[- ]?go['’]?s\b", False),
    "Asia": (r"\b(?:Asia|ASIA)\b", True), "Jennifer Holliday": (r"jennifer holliday", False),
    "Men at Work": (r"men at work", False), "Stray Cats": (r"\b(?:Stray Cats?|STRAY CATS?)\b", True),
    "The Human League": (r"human league", False),
    "Big Country": (r"\b(?:Big Country|BIG COUNTRY)\b", True), "Culture Club": (r"culture club", False),
    "Eurythmics": (r"eurythmics", False), "Men Without Hats": (r"men without hats", False),
    "Musical Youth": (r"musical youth", False), "Corey Hart": (r"corey hart", False),
    "Cyndi Lauper": (r"lauper", False),
    "Frankie Goes to Hollywood": (r"frankie goes to hollywood|\bFGTH\b", False),
    "Sheila E.": (r"sheila e\b|sheila escovedo", False),
    "The Judds": (r"\bthe judds\b|naomi judd|wynonna", False),
    "Freddie Jackson": (r"freddie jackson", False), "Julian Lennon": (r"julian lennon", False),
    "Katrina and the Waves": (r"katrina (?:and|&) the waves", False),
    "Sade": (r"(?<![Dd]e )\b(?:Sade|SADE)\b", True), "a-ha": (r"\ba-ha\b", False),
    "Bruce Hornsby and the Range": (r"hornsby", False), "Glass Tiger": (r"glass tiger", False),
    "Nu Shooz": (r"nu shooz", False), "Simply Red": (r"simply red", False), "Timbuk3": (r"timbuk ?3", False),
    "Breakfast Club": (r"breakfast club", False), "Cutting Crew": (r"cutting crew", False),
    "Jody Watley": (r"jody watley", False), "Swing Out Sister": (r"swing out sister", False),
    "Terence Trent D'Arby": (r"d['’]arby|sananda maitreya", False), "Rick Astley": (r"astley", False),
    "Take 6": (r"\b(?:Take (?:6|Six)|TAKE (?:6|SIX))\b", True), "Toni Childs": (r"toni childs", False),
    "Tracy Chapman": (r"tracy chapman", False), "Vanessa Williams": (r"vanessa williams", False),
    "Indigo Girls": (r"indigo girls", False), "Neneh Cherry": (r"neneh cherry", False),
    "Soul II Soul": (r"soul (?:ii|2) soul", False), "Tone Loc": (r"tone[- ]?l[oō]c", False),
}

MEMBERS = {
    "Dire Straits": [
        ("Mark Knopfler", ["Mark Knopfler"]),
        ("David Knopfler", ["David Knopfler"]),
        ("John Illsley", ["John Illsley"]),
        ("Pick Withers", ["Pick Withers", "David Withers", "David 'Pick' Withers"]),
    ],
    "The Knack": [
        ("Doug Fieger", ["Doug Fieger"]),
        ("Berton Averre", ["Berton Averre", "Burton Averre"]),
        ("Prescott Niles", ["Prescott Niles"]),
        ("Bruce Gary", ["Bruce Gary"]),
    ],
    "The Pretenders": [
        ("Chrissie Hynde", ["Chrissie Hynde", "Christine Ellen Hynde", "Chrissy Hynde"]),
        ("James Honeyman-Scott", ["James Honeyman-Scott", "Jimmy Scott"]),
        ("Pete Farndon", ["Pete Farndon", "Peter Farndon"]),
        ("Martin Chambers", ["Martin Chambers"]),
    ],
    "Adam and the Ants": [
        ("Adam Ant", ["Adam Ant", "Stuart Leslie Goddard", "Stuart Goddard"]),
        ("Marco Pirroni", ["Marco Pirroni"]),
        ("Terry Lee Miall", ["Terry Lee Miall"]),
        ("Chris Hughes", ["Chris Hughes", "Merrick", "Christopher Merrick Hughes", "Chris 'Merrick' Hughes"]),
        ("Gary Tibbs", ["Gary Tibbs", "Gary Brian Tibbs"]),
    ],
    "The Go-Go's": [
        ("Belinda Carlisle", ["Belinda Carlisle"]),
        ("Jane Wiedlin", ["Jane Wiedlin"]),
        ("Charlotte Caffey", ["Charlotte Caffey"]),
        ("Kathy Valentine", ["Kathy Valentine"]),
        ("Gina Schock", ["Gina Schock"]),
    ],
    "Asia": [
        ("John Wetton", ["John Wetton"]),
        ("Steve Howe", ["Steve Howe"]),
        ("Geoff Downes", ["Geoff Downes", "Geoffrey Downes"]),
        ("Carl Palmer", ["Carl Palmer"]),
    ],
    "Men at Work": [
        ("Colin Hay", ["Colin Hay"]),
        ("Ron Strykert", ["Ron Strykert"]),
        ("Greg Ham", ["Greg Ham"]),
        ("John Rees", ["John Rees"]),
        ("Jerry Speiser", ["Jerry Speiser"]),
    ],
    "Stray Cats": [
        ("Brian Setzer", ["Brian Setzer"]),
        ("Lee Rocker", ["Lee Rocker", "Leon Drucker"]),
        ("Slim Jim Phantom", ["Slim Jim Phantom", "James McDonnell", "Jim Phantom"]),
    ],
    "The Human League": [
        ("Philip Oakey", ["Philip Oakey", "Phil Oakey", "Phillip Oakey"]),
        ("Susan Ann Sulley", ["Susan Ann Sulley", "Susanne Sulley", "Susan Sulley"]),
        ("Joanne Catherall", ["Joanne Catherall"]),
        ("Philip Adrian Wright", ["Philip Adrian Wright", "Adrian Wright"]),
        ("Ian Burden", ["Ian Burden"]),
        ("Jo Callis", ["Jo Callis"]),
    ],
    "Big Country": [
        ("Stuart Adamson", ["Stuart Adamson", "William Stuart Adamson"]),
        ("Bruce Watson", ["Bruce Watson"]),
        ("Tony Butler", ["Tony Butler"]),
        ("Mark Brzezicki", ["Mark Brzezicki"]),
    ],
    "Culture Club": [
        ("Boy George", ["Boy George", "George Alan O'Dowd", "George O'Dowd"]),
        ("Jon Moss", ["Jon Moss"]),
        ("Mikey Craig", ["Mikey Craig", "Michael Emile Craig", "Michael Craig"]),
        ("Roy Hay", ["Roy Hay"]),
    ],
    "Eurythmics": [
        ("Annie Lennox", ["Annie Lennox", "Ann Lennox"]),
        ("Dave Stewart", ["Dave Stewart", "David Allan Stewart", "David A. Stewart"]),
    ],
    "Men Without Hats": [
        ("Ivan Doroschuk", ["Ivan Doroschuk"]),
        ("Stefan Doroschuk", ["Stefan Doroschuk"]),
        ("Colin Doroschuk", ["Colin Doroschuk"]),
        ("Allan McCarthy", ["Allan McCarthy"]),
    ],
    "Musical Youth": [
        ("Dennis Seaton", ["Dennis Seaton"]),
        ("Kelvin Grant", ["Kelvin Grant"]),
        ("Michael Grant", ["Michael Grant"]),
        ("Patrick Waite", ["Patrick Waite"]),
        ("Frederick Waite", ["Frederick Waite", "Freddie 'Junior' Waite", "Junior Waite"]),
    ],
    "Frankie Goes to Hollywood": [
        ("Holly Johnson", ["Holly Johnson", "William Johnson", "William 'Holly' Johnson"]),
        ("Paul Rutherford", ["Paul Rutherford"]),
        ("Peter Gill", ["Peter Gill", "Ped"]),
        ("Mark O'Toole", ["Mark O'Toole"]),
        ("Brian Nash", ["Brian Nash", "Nasher"]),
    ],
    "The Judds": [
        ("Naomi Judd", ["Naomi Judd", "Diana Ellen Judd", "Diana Ciminella"]),
        ("Wynonna Judd", ["Wynonna Judd", "Christina Claire Ciminella", "Wynonna Ellen Judd Moser", "Wynonna", "Christina Ciminella"]),
    ],
    "Katrina and the Waves": [
        ("Katrina Leskanich", ["Katrina Leskanich"]),
        ("Kimberley Rew", ["Kimberley Rew", "Kimberly Rew"]),
        ("Vince de la Cruz", ["Vince de la Cruz"]),
        ("Alex Cooper", ["Alex Cooper"]),
    ],
    "Sade": [
        ("Sade Adu", ["Sade Adu", "Helen Folasade Adu", "Folasade Adu", "Helen Adu"]),
        ("Stuart Matthewman", ["Stuart Matthewman"]),
        ("Paul S. Denman", ["Paul S. Denman", "Paul Denman", "Paul Spencer Denman"]),
        ("Andrew Hale", ["Andrew Hale"]),
    ],
    "a-ha": [
        ("Morten Harket", ["Morten Harket"]),
        ("Magne Furuholmen", ["Magne Furuholmen", "Mags", "Magne"]),
        ("Paul Waaktaar-Savoy", ["Paul Waaktaar-Savoy", "Pål Waaktaar", "Pål Waaktaar-Savoy", "Paul Waaktaar", "Pal Waaktaar"]),
    ],
    "Bruce Hornsby and the Range": [
        ("Bruce Hornsby", ["Bruce Hornsby"]),
        ("John Molo", ["John Molo"]),
        ("George Marinelli", ["George Marinelli"]),
        ("Joe Puerta", ["Joe Puerta", "Joseph Puerta"]),
        ("Peter Harris", ["Peter Harris"]),
        ("David Mansfield", ["David Mansfield"]),
    ],
    "Glass Tiger": [
        ("Alan Frew", ["Alan Frew"]),
        ("Sam Reid", ["Sam Reid"]),
        ("Al Connelly", ["Al Connelly"]),
        ("Wayne Parker", ["Wayne Parker"]),
        ("Michael Hanson", ["Michael Hanson"]),
    ],
    "Nu Shooz": [
        ("John Smith", ["John Smith"]),
        ("Valerie Day", ["Valerie Day"]),
    ],
    "Simply Red": [
        ("Mick Hucknall", ["Mick Hucknall", "Michael James Hucknall", "Michael Hucknall"]),
        ("Fritz McIntyre", ["Fritz McIntyre"]),
        ("Tony Bowers", ["Tony Bowers"]),
        ("Chris Joyce", ["Chris Joyce"]),
        ("Tim Kellett", ["Tim Kellett"]),
        ("Sylvan Richardson", ["Sylvan Richardson"]),
        ("Ian Kirkham", ["Ian Kirkham"]),
    ],
    "Timbuk3": [
        ("Pat MacDonald", ["Pat MacDonald"]),
        ("Barbara K. MacDonald", ["Barbara K. MacDonald", "Barbara Kooyman", "Barbara K"]),
    ],
    "Breakfast Club": [
        ("Dan Gilroy", ["Dan Gilroy"]),
        ("Ed Gilroy", ["Ed Gilroy", "Eddie Gilroy"]),
        ("Gary Burke", ["Gary Burke"]),
        ("Stephen Bray", ["Stephen Bray", "Steve Bray"]),
        ("Paul Kauk", ["Paul Kauk"]),
        ("Randy Jackson", ["Randy Jackson"]),
    ],
    "Cutting Crew": [
        ("Nick Van Eede", ["Nick Van Eede", "Nicholas Eede"]),
        ("Kevin MacMichael", ["Kevin MacMichael", "Kevin Scott MacMichael"]),
        ("Colin Farley", ["Colin Farley"]),
        ("Martin Beedle", ["Martin Beedle", "Frosty Beedle", "Martin 'Frosty' Beedle", "Martin Philip Beedle"]),
        ("Tony Moore", ["Tony Moore"]),
    ],
    "Swing Out Sister": [
        ("Corinne Drewery", ["Corinne Drewery", "Corrine Drewery"]),
        ("Andy Connell", ["Andy Connell", "Andrew Connell"]),
        ("Martin Jackson", ["Martin Jackson"]),
    ],
    "Take 6": [
        ("Claude McKnight", ["Claude McKnight", "Claude V. McKnight III"]),
        ("Mark Kibble", ["Mark Kibble"]),
        ("Mervyn Warren", ["Mervyn Warren"]),
        ("Cedric Dent", ["Cedric Dent"]),
        ("David Thomas", ["David Thomas", "Dave Thomas"]),
        ("Alvin Chea", ["Alvin Chea"]),
    ],
    "Indigo Girls": [
        ("Amy Ray", ["Amy Ray"]),
        ("Emily Saliers", ["Emily Saliers"]),
    ],
    "Soul II Soul": [
        ("Jazzie B", ["Jazzie B", "Trevor Beresford Romeo", "Beresford Romeo"]),
        ("Nellee Hooper", ["Nellee Hooper", "Paul Andrew Hooper", "Nellie Hooper"]),
        ("Caron Wheeler", ["Caron Wheeler", "Caron Melina Wheeler"]),
        ("Daddae Harvey", ["Daddae Harvey", "Daddae", "Philip Harvey", "Philip 'Daddae' Harvey"]),
        ("Aitch B", ["Aitch B", "Aitch Bee", "DJ Aitch B"]),
        ("Simon Law", ["Simon Law"]),
    ],
    "Rickie Lee Jones": [
        ("Rickie Lee Jones", ["Rickie Lee Jones"]),
    ],
    "Christopher Cross": [
        ("Christopher Cross", ["Christopher Cross", "Chris Cross", "Christopher Charles Geppert"]),
    ],
    "Irene Cara": [
        ("Irene Cara", ["Irene Cara", "Irene Cara Escalera"]),
    ],
    "Robbie Dupree": [
        ("Robbie Dupree", ["Robbie Dupree", "Robert Dupuis"]),
    ],
    "Amy Holland": [
        ("Amy Holland", ["Amy Holland"]),
    ],
    "Sheena Easton": [
        ("Sheena Easton", ["Sheena Easton", "Sheena Shirley Orr"]),
    ],
    "James Ingram": [
        ("James Ingram", ["James Ingram"]),
    ],
    "Luther Vandross": [
        ("Luther Vandross", ["Luther Vandross", "Luther Ronzoni Vandross"]),
    ],
    "Jennifer Holliday": [
        ("Jennifer Holliday", ["Jennifer Holliday"]),
    ],
    "Cyndi Lauper": [
        ("Cyndi Lauper", ["Cyndi Lauper", "Cindy Lauper", "Cynthia Lauper"]),
    ],
    "Sheila E.": [
        ("Sheila E.", ["Sheila E.", "Sheila Escovedo", "Sheila Cecilia Escovedo"]),
    ],
    "Corey Hart": [
        ("Corey Hart", ["Corey Hart", "Corey Mitchell Hart"]),
    ],
    "Freddie Jackson": [
        ("Freddie Jackson", ["Freddie Jackson"]),
    ],
    "Julian Lennon": [
        ("Julian Lennon", ["Julian Lennon", "John Charles Julian Lennon"]),
    ],
    "Jody Watley": [
        ("Jody Watley", ["Jody Watley"]),
    ],
    "Terence Trent D'Arby": [
        ("Terence Trent D'Arby", ["Terence Trent D'Arby", "Sananda Maitreya", "Terence Trent Howard", "Trent D'Arby", "Francesco Maitreya"]),
    ],
    "Tracy Chapman": [
        ("Tracy Chapman", ["Tracy Chapman"]),
    ],
    "Rick Astley": [
        ("Rick Astley", ["Rick Astley", "Richard Paul Astley"]),
    ],
    "Toni Childs": [
        ("Toni Childs", ["Toni Childs"]),
    ],
    "Vanessa Williams": [
        ("Vanessa Williams", ["Vanessa Williams", "Vanessa Lynn Williams"]),
    ],
    "Neneh Cherry": [
        ("Neneh Cherry", ["Neneh Cherry", "Neneh Mariann Karlsson"]),
    ],
    "Tone Loc": [
        ("Tone Loc", ["Tone Loc", "Anthony Terrell Smith", "Anthony Smith"]),
    ],
}

# One-word nicknames distinctive enough to count on their own (matched with
# their capital letter)
ONE_WORD_OK = {"Merrick", "Nasher", "Wynonna", "Mags", "Daddae", "Magne", "Ped"}
# A member's first name that journalists use on its own, counted (like a
# surname) only after the member has been named in full in the same record
FIRST_NAME_FOLLOW = {
    "Cyndi Lauper": "Cyndi", "Neneh Cherry": "Neneh", "Luther Vandross": "Luther", "Sheena Easton": "Sheena",
    "Chrissie Hynde": "Chrissie", "Magne Furuholmen": "Magne", "Morten Harket": "Morten",
    "Wynonna Judd": "Wynonna", "Naomi Judd": "Naomi", "Holly Johnson": "Holly", "Rickie Lee Jones": "Rickie",
    "Tracy Chapman": "Tracy", "Toni Childs": "Toni", "Vanessa Williams": "Vanessa", "Jody Watley": "Jody",
    "Irene Cara": "Irene", "Corey Hart": "Corey", "Freddie Jackson": "Freddie", "Julian Lennon": "Julian",
    "Rick Astley": "Rick", "Amy Holland": "Amy", "Robbie Dupree": "Robbie", "Jennifer Holliday": "Jennifer",
    "Sheila E.": "Sheila", "Terence Trent D'Arby": "Terence", "Jazzie B": "Jazzie", "Annie Lennox": "Annie",
    "Belinda Carlisle": "Belinda", "Boy George": "George",
}
# surname-like tokens never counted on their own
NO_FOLLOW = {"B", "E.", "Q", "Jr.", "Jr", "Sr.", "Sr", "III", "II", "Loc", "Ant"}
# capitalised words that can stand before a surname without being someone's
# first name ("When Hart arrived", "Mr. Hart")
NOT_A_FIRST_NAME = {
    "The", "A", "An", "And", "But", "As", "When", "While", "If", "So", "Then", "Says", "Said", "Asked",
    "Mr.", "Ms.", "Mrs.", "Dr.", "Mr", "Ms", "Mrs", "Dr", "Like", "With", "For", "From", "By", "After",
    "Before", "Since", "Although", "Though", "Because", "Meanwhile", "Now", "Today", "Also", "Even", "Only",
    "Just", "In", "On", "At", "Of", "To", "Recalls", "Explains", "Adds", "Notes", "Admits", "Laughs",
    "Told", "Tells", "Here", "There", "Yes", "No", "Well", "Oh", "Sir", "Q", "A:", "Q:"}

# ═════════════════════════════════════════════════════════════════════════════
# MEDIA PAGES
# ═════════════════════════════════════════════════════════════════════════════
MEDIA_SITE = re.compile(
    r"^https?://(?:[\w-]+\.)*(?:youtube\.com|youtu\.be|vimeo\.com|dailymotion\.com|soundcloud\.com|"
    r"podcasts\.apple\.com|open\.spotify\.com|iheart\.com|podbean\.com|anchor\.fm|libsyn\.com|"
    r"buzzsprout\.com|podomatic\.com|mixcloud\.com|art19\.com|acast\.com)/"
    r"|bbc\.co\.uk/(?:sounds/play|programmes|iplayer)|pbs\.org/video|cbc\.ca/(?:listen|player)"
    r"|abc\.net\.au/listen|/(?:podcasts?|episodes?)/", re.I)
MEDIA_SITE_NOT = re.compile(r"iheart\.com/artist/", re.I)       # artist radio stations, not interviews
INTERVIEW_WORD = re.compile(r"\b(?:interviews?|interviewed|sessions?|talks?|talking|in conversation|"
                            r"conversation|chats?|Q ?& ?A|discuss(?:es|ed)?|speaks?|guest|oral history)\b", re.I)
MEDIA_WORD = re.compile(r"\b(?:videos?|audio|podcasts?|episodes?|radio|listen|watch|tv|television|"
                        r"broadcast|sounds|webcast|livestream|stream)\b", re.I)
NOT_ABOUT_ACT = {"act name used in another sense", "namesake: another person", "index or listing page",
                 "never names the act or a member"}


# ─────────────────────────────────────────────────────────────────────────────
def blank(v):
    return v is None or (isinstance(v, float) and v != v) or \
        str(v).strip().lower() in ("", "nan", "none", "nat")


def text_of(v):
    return "" if blank(v) else str(v)


def url_hash(url):
    return hashlib.md5(str(url).encode("utf-8")).hexdigest()[:10]


def in_turn_set(rec):
    return str(rec.get("Date Confidence", "")).strip().lower() == "high" and blank(rec.get("Turn Set Exclusion"))


def parse_exclusions():
    out = {}
    for line in EXCLUSIONS_TSV.strip().splitlines():
        rid, code, h, act = line.split("\t")
        out[int(rid)] = (REASONS[code], h, act)
    return out


# --- relevance ---------------------------------------------------------------
SPACE = re.compile(r"[\s  -​  　]")   # one character each, so positions hold


def flat(s):
    return SPACE.sub(" ", s)


def name_pattern(name):
    p = re.escape(name)
    p = p.replace("\\ ", " ").replace(" ", r"[\s‐‑–-]+")
    p = p.replace("'", "['’]").replace("\\-", "[-‐‑–]")
    return p


def surname_of(member):
    toks = [t for t in re.sub(r"\s*['’][^'’]+['’]\s*", " ", member).split()
            if t not in {"Jr.", "Jr", "Sr.", "Sr", "III", "II"}]
    if len(toks) < 2:
        return None
    s = toks[-1]
    return None if s in NO_FOLLOW or len(s.strip(".")) < 3 else s


class ActNames:
    def __init__(self, act):
        pat, cs = ACT_NAME[act]
        self.main = [re.compile(pat, 0 if cs else re.I)]
        multi, one = [], []
        self.intro = []          # (member, regex that shows the member has been named in full)
        follow = {}              # token -> members it can stand for
        for member, names in MEMBERS.get(act, []):
            m_multi = [n for n in names if " " in n.strip()]
            m_one = [n for n in names if n in ONE_WORD_OK]
            multi += m_multi
            one += m_one
            parts = [name_pattern(n) for n in m_multi]
            if parts:
                self.intro.append((member, re.compile(r"(?<![\w-])(?:" + "|".join(parts) + r")(?![\w-])", re.I)))
            for tok in filter(None, [surname_of(member), FIRST_NAME_FOLLOW.get(member)]):
                follow.setdefault(tok, set()).add(member)
        multi = sorted(set(multi), key=len, reverse=True)
        if multi:
            self.main.append(re.compile(r"(?<![\w-])(?:" + "|".join(name_pattern(n) for n in multi) + r")(?![\w-])", re.I))
        if one:
            self.main.append(re.compile(r"(?<![\w-])(?:" + "|".join(sorted(set(one))) + r")(?![\w-])"))
        self.follow = follow
        self.follow_rx = None
        if follow:
            alts = []
            for tok in sorted(follow, key=len, reverse=True):
                alts += [name_pattern(tok), name_pattern(tok.upper())]
            # not glued to a word, and not followed by another capitalised word or an initial
            # ("Jazzie Q", "Corey Feldman")
            self.follow_rx = re.compile(r"(?<![\w'’-])(" + "|".join(alts) + r")(?![\w-])(?!\s+[A-Z][\w.'’-]*)")

    def count(self, title, text):
        ti, tx = flat(title), flat(text)
        in_title = any(rx.search(ti) for rx in self.main)
        taken = []
        for rx in self.main:
            for m in rx.finditer(tx):
                if not any(m.start() < e and m.end() > s for s, e in taken):
                    taken.append((m.start(), m.end()))
        introduced = {mem for mem, rx in self.intro if rx.search(ti) or rx.search(tx)}
        if self.follow_rx and introduced:
            for m in self.follow_rx.finditer(tx):
                s, e = m.start(), m.end()
                if any(s < te and e > ts for ts, te in taken):
                    continue
                tok = m.group(1)
                key = next((k for k in self.follow if k == tok or k.upper() == tok), None)
                if key is None or not (self.follow[key] & introduced):
                    continue
                before = re.search(r"([A-Z][\w'’.-]*)\s+$", tx[max(0, s - 40):s])
                if before and before.group(1) not in NOT_A_FIRST_NAME:
                    continue           # "Mickey Hart", "Doris Day": another person
                taken.append((s, e))
        n = len(taken)
        first = min(s for s, _ in taken) if taken else -1
        if in_title:
            label = "named in title"
        elif n >= 2:
            label = "named 2+ times"
        elif n == 1:
            label = "named once"
        else:
            label = "not found"
        return in_title, n, first, label


def media_page(url, title, text):
    u = str(url)
    if MEDIA_SITE.search(u) and not MEDIA_SITE_NOT.search(u):
        return "audio/video site"
    t = text.strip()
    if len(t) < SHORT_PAGE:
        head = title + " " + t[:600]
        if INTERVIEW_WORD.search(head) and MEDIA_WORD.search(head):
            return "short audio/video page"
    return ""


def excel_safe(frame, text_cap):
    out = frame.copy()
    if "Extracted Text" in out.columns:
        out["Extracted Text"] = out["Extracted Text"].map(lambda v: text_of(v)[:text_cap])
    for c in out.columns:
        if pd.api.types.is_object_dtype(out[c]) or pd.api.types.is_string_dtype(out[c]):
            out[c] = out[c].map(lambda v: ILLEGAL_CHARACTERS_RE.sub("", v) if isinstance(v, str) else v)
    return out


def stop(msg):
    print("\nSTOPPED, nothing was written:\n  " + msg.replace("\n", "\n  "))
    sys.exit(1)


# ─────────────────────────────────────────────────────────────────────────────
def load_input(excl):
    tried = []
    for name in IN_JSON_CANDIDATES:
        path = os.path.join(DIR, name)
        if not os.path.exists(path):
            tried.append(f"{name}: not found")
            continue
        with open(path, encoding="utf-8") as f:
            recs = json.load(f)
        problems = []
        ids = [r.get("Record ID") for r in recs]
        if len(recs) != EXPECTED_RECORDS:
            problems.append(f"{len(recs)} records, expected {EXPECTED_RECORDS}")
        if [int(i) for i in ids if not blank(i)] != list(range(1, len(recs) + 1)):
            problems.append("Record IDs are not 1, 2, 3 ... in order")
        ts = sum(in_turn_set(r) for r in recs)
        if ts != EXPECTED_TURN_SET_BEFORE:
            problems.append(f"turn set is {ts}, expected {EXPECTED_TURN_SET_BEFORE}")
        if not problems:
            by_id = {int(r["Record ID"]): r for r in recs}
            bad = []
            for rid, (_, h, act) in excl.items():
                r = by_id.get(rid)
                if r is None:
                    bad.append(f"{rid}: missing")
                elif url_hash(r.get("URL")) != h or str(r.get("Focal Act")) != act:
                    bad.append(f"{rid}: URL or act differs ({r.get('Focal Act')})")
                elif not in_turn_set(r):
                    bad.append(f"{rid}: not in the turn set")
            if bad:
                problems.append(f"{len(bad)} reviewed records do not match, e.g. " + "; ".join(bad[:5]))
        if problems:
            tried.append(f"{name}: " + "; ".join(problems))
            continue
        return path, recs
    stop("No CLEAN6 file in the folder matches the version the review was done on.\n" + "\n".join(tried))


def main():
    print(VERSION)
    excl = parse_exclusions()
    if len(excl) != 975 or Counter(r for r, _, _ in excl.values()) != Counter(
            {REASONS["A"]: 632, REASONS["N"]: 300, REASONS["U"]: 24, REASONS["P"]: 19}):
        stop("The exclusion table inside this script is damaged; download the script again.")
    path, recs = load_input(excl)
    print(f"Read {os.path.basename(path)}: {len(recs)} records, turn set {EXPECTED_TURN_SET_BEFORE}")
    existing = {text_of(r.get("Turn Set Exclusion")).strip() for r in recs} - {""}
    unknown = existing - OLD_REASONS
    if unknown:
        stop(f"Unexpected exclusion reasons already in the file: {sorted(unknown)}")
    if any(c in recs[0] for c in NEW_COLUMNS):
        stop("The input already has CLEAN7 columns; run this on CLEAN6, not on CLEAN7.")
    before = json.dumps(recs, ensure_ascii=False, default=str, sort_keys=True)

    # 1. exclusions
    for r in recs:
        rid = int(r["Record ID"])
        old = text_of(r.get("Turn Set Exclusion")).strip()
        if rid in excl:
            r["Turn Set Exclusion"] = excl[rid][0]
            r["Turn Set Exclusion Source"] = SOURCE_NEW
            r["Turn Set Exclusion Note"] = NOTES.get(rid, "")
        else:
            r["Turn Set Exclusion Source"] = SOURCE_OLD if old else ""
            r["Turn Set Exclusion Note"] = ""
        r["In Turn Set"] = "Yes" if in_turn_set(r) else "No"

    # 2 and 3. relevance and media page, for every record
    names = {}
    for r in recs:
        act = str(r.get("Focal Act"))
        if act not in ACT_NAME:
            stop(f"Record {r.get('Record ID')}: act '{act}' has no name pattern in this script.")
        if act not in names:
            names[act] = ActNames(act)
        title, text = text_of(r.get("Title / Headline")), text_of(r.get("Extracted Text"))
        in_title, n, first, label = names[act].count(title, text)
        r["Relevance"] = label
        r["Relevance Mentions"] = n
        r["Named In Title"] = "Yes" if in_title else "No"
        r["First Mention At"] = first
        r["Media Page"] = media_page(r.get("URL"), title, text)

    # checks: only the intended changes
    after = [{k: v for k, v in r.items() if k not in NEW_COLUMNS} for r in recs]
    for r, a in zip(recs, after):
        if int(r["Record ID"]) in excl:
            a["Turn Set Exclusion"] = None        # the one field allowed to change
    chk_before = json.loads(before)
    for r in chk_before:
        if int(r["Record ID"]) in excl:
            r["Turn Set Exclusion"] = None
    if json.dumps(after, ensure_ascii=False, default=str, sort_keys=True) != \
            json.dumps(chk_before, ensure_ascii=False, default=str, sort_keys=True):
        stop("A value other than Turn Set Exclusion changed. This is a bug; please report it with the run's output.")
    ts_after = sum(r["In Turn Set"] == "Yes" for r in recs)
    if ts_after != EXPECTED_TURN_SET_AFTER:
        stop(f"The turn set came out at {ts_after}, expected {EXPECTED_TURN_SET_AFTER}.")

    # summaries
    df = pd.DataFrame(recs)
    ts = df["In Turn Set"] == "Yes"
    media_rows = df[(df["Media Page"] != "") & ~df["Turn Set Exclusion"].map(text_of).isin(NOT_ABOUT_ACT)]
    leads = df[df["Record ID"].astype(int).isin(LEADS)].copy()
    summary = summarise(df, ts)

    if DRY_RUN:
        report(df, ts, media_rows, summary, wrote=False)
        return

    # write JSON (values exactly as read, plus the new keys)
    with open(OUT_JSON, "w", encoding="utf-8") as f:
        json.dump(recs, f, ensure_ascii=False, default=str)
    # workbook
    rejected = None
    for name in IN_XLSX_CANDIDATES:
        p = os.path.join(DIR, name)
        if os.path.exists(p):
            with pd.ExcelFile(p) as xf:
                if "Rejected" in xf.sheet_names:
                    rejected = pd.read_excel(xf, sheet_name="Rejected")
            break
    high = df["Date Confidence"].astype(str).str.strip().str.lower() == "high"
    ex = df[high & (df["Turn Set Exclusion"].map(text_of) != "")].copy()
    ex_view = ex[["Record ID", "Focal Act", "Grammy Year", "Turn Set Exclusion", "Turn Set Exclusion Source",
                  "Turn Set Exclusion Note", "Page Type", "Temporal Class", "Title / Headline", "URL",
                  "Relevance", "Relevance Mentions"]].copy()
    ex_view["Opening text"] = ex["Extracted Text"].map(lambda v: text_of(v)[:300])
    ex_view = ex_view.sort_values(["Turn Set Exclusion Source", "Turn Set Exclusion", "Focal Act", "Record ID"])
    with pd.ExcelWriter(OUT_XLSX, engine="openpyxl") as xw:
        readme(df, ts, summary).to_excel(xw, sheet_name="Read me", index=False, header=False)
        excel_safe(df, EXCEL_CELL_CAP).to_excel(xw, sheet_name="Corpus", index=False)
        if rejected is not None:
            excel_safe(rejected, REJECTED_TEXT_CAP).to_excel(xw, sheet_name="Rejected", index=False)
        excel_safe(ex_view, EXCEL_CELL_CAP).to_excel(xw, sheet_name="Turn set exclusions", index=False)
        summary.to_excel(xw, sheet_name="Turn set by act", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2" if ws.title != "Read me" else None
        xw.book["Read me"].column_dimensions["A"].width = 110
    write_media(df, media_rows, leads)
    # re-read and confirm
    with open(OUT_JSON, encoding="utf-8") as f:
        back = json.load(f)
    if len(back) != EXPECTED_RECORDS or sum(r.get("In Turn Set") == "Yes" for r in back) != EXPECTED_TURN_SET_AFTER:
        stop("The written JSON did not read back correctly.")
    report(df, ts, media_rows, summary, wrote=True)


def summarise(df, ts):
    rows = []
    for act, g in df.groupby("Focal Act", sort=False):
        gts = g[g["In Turn Set"] == "Yes"]
        before = int(((g["In Turn Set"] == "Yes") | g["Turn Set Exclusion Source"].eq(SOURCE_NEW)).sum())
        row = {"Focal Act": act, "Grammy Year": g["Grammy Year"].iloc[0], "Turn set before": before,
               "Excluded by the review": int(g["Turn Set Exclusion Source"].eq(SOURCE_NEW).sum()),
               "Turn set now": len(gts)}
        for lab in ["named in title", "named 2+ times", "named once", "not found"]:
            row[lab.capitalize()] = int((gts["Relevance"] == lab).sum())
        row["Media pages"] = int((gts["Media Page"] != "").sum())
        rows.append(row)
    out = pd.DataFrame(rows).sort_values(["Grammy Year", "Focal Act"])
    return out


def readme(df, ts, summary):
    t = df[ts]
    rc = t["Relevance"].value_counts()
    lines = [
        "CLEAN7  (" + VERSION + ")",
        "",
        "What changed from CLEAN6",
        "975 turn-set records were excluded after the namesake review of 4 October 2026. Nothing was deleted, "
        "renumbered or re-dated; the only existing value that changed is Turn Set Exclusion for those 975 records.",
        "Turn set: 10,291 records before, 9,316 now (high date confidence and a blank Turn Set Exclusion).",
        "",
        "New columns",
        "Turn Set Exclusion Source: where each exclusion came from (the finalise_v7 rules of 28 September, or the "
        "namesake review of 4 October).",
        "Turn Set Exclusion Note: the note written during the review, where there is one, and how the two "
        "undecided rows were settled from the stored text.",
        "In Turn Set: Yes or No, the turn-set rule spelled out.",
        "Relevance: named in title / named 2+ times / named once / not found. A flag only. Records named once are "
        "almost always passing mentions; records named 2+ times include roundups and lists as well as articles "
        "about the act, so the turn build decides which parts are about the act.",
        "Relevance Mentions, Named In Title, First Mention At: the evidence behind Relevance. Mentions count the act's "
        "name and nomination-era members' names and aliases, and a member's surname (or a distinctive first name) "
        "after the member has been named in full.",
        "Media Page: audio/video site, or short audio/video page. These pages hold an interview as audio or video, "
        "not text; they are listed in harrison_v7_MEDIA_PAGES.xlsx for the video work.",
        "",
        "Turn set by relevance",
    ] + [f"{k}: {int(rc.get(k, 0)):,}" for k in ["named in title", "named 2+ times", "named once", "not found"]] + [
        f"Media pages in the turn set: {int((t['Media Page'] != '').sum()):,}",
        "",
        "Sheets",
        "Corpus: every record. Rejected: as in CLEAN6. Turn set exclusions: every high-confidence record that is "
        "excluded, with its reason, source and note. Turn set by act: counts before and after, with relevance.",
    ]
    return pd.DataFrame({"A": lines})


def write_media(df, media_rows, leads):
    cols = ["Record ID", "Focal Act", "Grammy Year", "Media Page", "Title / Headline", "Outlet / Source Name",
            "URL", "Publication Date", "Date Confidence", "Temporal Class", "In Turn Set", "Turn Set Exclusion",
            "Relevance", "Text Length"]
    out = pd.concat([media_rows[cols], leads[~leads["Record ID"].isin(media_rows["Record ID"])][cols]],
                    ignore_index=True)
    out.insert(4, "Lead", out["Record ID"].astype(int).map(LEADS).fillna(""))
    opening = dict(zip(df["Record ID"], df["Extracted Text"].map(lambda v: text_of(v)[:300])))
    out["Opening text"] = out["Record ID"].map(opening).fillna("")
    out = out.sort_values(["Grammy Year", "Focal Act", "Publication Date"], na_position="last")
    with pd.ExcelWriter(OUT_MEDIA, engine="openpyxl") as xw:
        pd.DataFrame({"A": [
            "Media pages: interviews held as audio or video",
            "",
            "Each row is a corpus record whose page is a podcast, video or radio player, or a short page announcing "
            "an audio or video interview. Their text is too thin for the turn build, but the interview itself may be "
            "worth finding for the video work.",
            "Media Page says how the page was recognised. Lead holds notes on pages kept for another reason.",
            "Records excluded as namesakes, other senses of the act's name, index pages or never naming the act are "
            "left out. Records with lower date confidence are kept, since the interview may still be useful.",
            f"Rows: {len(out)}",
        ]}).to_excel(xw, sheet_name="Read me", index=False, header=False)
        excel_safe(out, EXCEL_CELL_CAP).to_excel(xw, sheet_name="Media pages", index=False)
        xw.book["Read me"].column_dimensions["A"].width = 110
        xw.book["Media pages"].freeze_panes = "A2"


def report(df, ts, media_rows, summary, wrote):
    t = df[ts]
    print(f"\n{'=' * 66}")
    print(f"  records                       : {len(df)}")
    print(f"  turn set before               : {EXPECTED_TURN_SET_BEFORE}")
    new = df[df["Turn Set Exclusion Source"] == SOURCE_NEW]
    print(f"  excluded by the review        : {len(new)}")
    for reason, n in new["Turn Set Exclusion"].value_counts().items():
        print(f"      {reason:32} {n}")
    print(f"  TURN SET NOW                  : {int(ts.sum())}")
    print("\n  Turn set by relevance:")
    for k in ["named in title", "named 2+ times", "named once", "not found"]:
        print(f"      {k:18} {int((t['Relevance'] == k).sum())}")
    print(f"  media pages in the turn set   : {int((t['Media Page'] != '').sum())}")
    print(f"  rows in the media list        : {len(media_rows)}  (+ {len(LEADS)} other lead)")
    print("\n  Acts losing the most records:")
    s = summary.assign(lost=summary["Excluded by the review"]).sort_values("lost", ascending=False).head(6)
    for _, r in s.iterrows():
        print(f"      {r['Focal Act']:28} {r['Turn set before']:>4} -> {r['Turn set now']:>4}")
    print("\n  Thinnest acts now:")
    for _, r in summary.sort_values("Turn set now").head(5).iterrows():
        print(f"      {r['Focal Act']:28} {r['Turn set now']:>4}")
    if wrote:
        print(f"\nWrote:\n  {OUT_JSON}\n  {OUT_XLSX}\n  {OUT_MEDIA}")
    else:
        print("\nDRY RUN: every check passed; nothing was written.")


if __name__ == "__main__":
    main()
