from pathlib import Path

randomlist = []

for l in list(Path(r"..\TestReport\2017-07-26_154721\TC4").rglob('*.random')):
    p = str(list(l.parts)[-1:])
    # l = str(l)
    p = p[:p.find(".random")]
    # print("p is %s" % p)
    randomlist.append(p.split("_"))

    #print(l)
    #randomlist.append()
#print(randomlist)

r = list(randomlist[0])


print(randomlist[0][0][-1:])

#file = Path(r".") / Path("test.random")

#file.touch()
