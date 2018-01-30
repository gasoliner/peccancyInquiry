class AttrDisplay:
    def gatherAttrs(self):
        return ",".join("{}={}"
                        .format(k, getattr(self, k))
                        for k in self.__dict__.keys())
        # attrs = []
        # for k in self.__dict__.keys():
        #     item = "{}={}".format(k, getattr(self, k))
        #     attrs.append(item)
        # return attrs
        # for k in self.__dict__.keys():
        #     attrs.append(str(k) + "=" + str(self.__dict__[k]))
        # return ",".join(attrs) if len(attrs) else 'no attr'

    def __str__(self):
        return "[{}:{}]".format(self.__class__.__name__, self.gatherAttrs())

class Properties(object):

    def __init__(self, fileName):
        self.fileName = fileName
        self.properties = {}

    def __getDict(self,strName,dictName,value):

        if(strName.find('.')>0):
            k = strName.split('.')[0]
            dictName.setdefault(k,{})
            return self.__getDict(strName[len(k)+1:],dictName[k],value)
        else:
            dictName[strName] = value
            return
    def getProperties(self):
        try:
            pro_file = open(self.fileName, 'Ur')
            for line in pro_file.readlines():
                line = line.strip().replace('\n', '')
                if line.find("#")!=-1:
                    line=line[0:line.find('#')]
                if line.find('=') > 0:
                    strs = line.split('=')
                    strs[1]= line[len(strs[0])+1:]
                    self.__getDict(strs[0].strip(),self.properties,strs[1].strip())
        except Exception, e:
            raise e
        else:
            pro_file.close()
        return self.properties

class CarInfo(AttrDisplay):

    def __init__(self, name, carnum, enginnum, kind):
        self.name = name
        self.carnum = carnum
        self.enginnum = enginnum
        self.kind = kind
