import os
import win32com.client
import pint
import numpy as np

u = pint.UnitRegistry()
UNITS = {'lb_in_F':1,'lb_ft_F':2,'kip_in_F':3,'kip_ft_F':4,
         'kN_mm_C':5,'kN_m_C':6,'kgf_mm_C':7,'kgf_m_C':8,
         'N_mm_C':9,'N_m_C':10,'Ton_mm_C':11,'Ton_m_C':12,
         'kN_cm_C':13,'kgf_cm_C':14,'N_cm_C':15,'Ton_cm_C':16}
MATERIALS = {'Steel': 1, 'Concrete':2, 'Nodesign':3, 'Aluminum':4, 'ColdFormed':5, 'Rebar':6, 'Tendom':7}
LOAD_PATTERN_TYPE = {'LTYPE_DEAD':1, 'LTYPE_SUPERDEAD':2,
                     'LTYPE_LIVE':3,'LTYPE_REDUCELIVE':4,'LTYPE_QUAKE':5,'LTYPE_WIND':6,
                     'LTYPE_SNOW':7,'LTYPE_OTHER':8,'LTYPE_MOVE':9,'LTYPE_TEMPERATURE':10,
                     'LTYPE_ROOFLIVE':11,'LTYPE_NOTIONAL':12,'LTYPE_PATTERNLIVE':13,'LTYPE_WAVE':4,
                     'LTYPE_BRAKING':15,'LTYPE_CENTRIFUGAL':16,'LTYPE_FRICTION':17,'LTYPE_ICE':18,
                     'LTYPE_WINDONLIVELOAD':19,'LTYPE_HORIZONTALEARTHPRESSURE':20,'LTYPE_VERTICALEARTHPRESSURE':21,
                     'LTYPE_EARTHSURCHARGE':22,'LTYPE_DOWNDRAG':23,'LTYPE_VEHICLECOLLISION':24,'LTYPE_VESSELCOLLISION':25,
                     'LTYPE_TEMPERATUREGRADIENT':26,'LTYPE_SETTLEMENT':27,'LTYPE_SHRINKAGE':28,'LTYPE_CREEP':29,'LTYPE_WATERLOADPRESSURE':30,
                     'LTYPE_LIVELOADSURCHARGE':31,'LTYPE_LOCKEDINFORCES':32,'LTYPE_PEDESTRIANLL':33,'LTYPE_PRESTRESS':34,'LTYPE_HYPERSTATIC':35,
                     'LTYPE_BOUYANCY':36,'LTYPE_STREAMFLOW':37,'LTYPE_IMPACT':38,'LTYPE_CONSTRUCTION':39}
class SAPy2000:
    def __init__(self):
        print('----- SAP2000 Starting ----- ')
        self.SapObject = win32com.client.Dispatch("Sap2000v16.SapObject")

        self.SapObject.ApplicationStart()

        self.SapModel = self.SapObject.SapModel
        print('----- Started -----')

    def newModel(self):
        #initialize model
        self.SapModel.InitializeNewModel()

        #create new blank model
        ret = self.SapModel.File.NewBlank()
        print('----- New blank -----')

    def setUnits(self, units='Ton_m_C'):
        ret = self.SapModel.SetPresentUnits(UNITS[units])
        print(f'Units set to: {units}')

    def defineConcrete(self, fck):
        fck = fck*u.MPa
        Eci = 5600*(fck.magnitude)**(1/2)*u.MPa
        Ecs = 0.85*Eci
        coef_poisson = 0.2
        coef_term = 10**(-5)/u.degC
        densidade = 2.5*u.tf/u.m**3

        Name = f'C{fck.magnitude}'
        ret = self.SapModel.PropMaterial.SetMaterial(Name=Name, MatType=MATERIALS['Concrete'])
        ret = self.SapModel.PropMaterial.SetMPIsotropic(Name=Name,
                                                   e=Eci.to('tf/m**2').magnitude,
                                                   u=coef_poisson,
                                                   a=coef_term.magnitude)
        ret = self.SapModel.PropMaterial.SetOConcrete_1(Name=Name,
                                                   fc=fck.to('tf/m**2').magnitude,
                                                   IsLightweight=False,
                                                   fcsfactor=0,
                                                   SSType=1,
                                                   SSHysType=2,
                                                   StrainAtfc=0.0022,
                                                   StrainUltimate=0.0052,
                                                   FinalSlope=-0.1)
        ret = self.SapModel.PropMaterial.SetWeightAndMass(Name=Name,
                                                     MyOption=1,
                                                    Value=densidade.magnitude)
        print(f'Material CONCRETE fck={fck} created!')

    def defineSolidProp(self, matProp):
        name = f'Solid_{matProp}'

        ret = self.SapModel.PropSolid.SetProp(name,matProp,0, 0, 0, True)
        assert ret == 0

        print(f'{self.SapModel.PropSolid.GetNameList()}')

    def create_regularSolid(self, length, width, height, solidProp):
        x = np.array([0, length])
        y = np.array([0, width])
        z = np.array([0, height])

        xx, yy, zz = np.meshgrid(x,y,z)
        xx, yy, zz = np.ravel(xx), np.ravel(yy), np.ravel(zz)

        ret = self.SapModel.SolidObj.AddByCoord(xx,yy,zz, f'Solid_{length}x{width}x{height}', solidProp)
        print(ret)
        ret = self.SapModel.View.RefreshView(0, False)
