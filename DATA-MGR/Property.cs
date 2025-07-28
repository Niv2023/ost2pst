
using System.Text;

namespace ost2pst
{
    // Enums and classes used in property handling
    // Enum names according to <MS-PST>

    public class Property
    {
        public EpropertyId id;
        public EpropertyType type;
        public HNID hnid;
        //        private PCBTH _pcbth;
        private byte[] _bytes;
        public PCBTH PCBTH { get { return new PCBTH() { wPropId = id, wPropType = type, dwValueHnid = hnid }; } }
        public byte[] data { get { return _bytes; } }
        public Property(PCBTH pc)
        {
            id = pc.wPropId;
            type = pc.wPropType;
            hnid = pc.dwValueHnid;
            switch (type)
            {
                case EpropertyType.PtypInteger16:
                    _bytes = new byte[2];
                    FM.TypeToArray<UInt16>(ref _bytes, (UInt16)hnid.dwValue);
                    break;
                case EpropertyType.PtypInteger32:
                    _bytes = new byte[4];
                    FM.TypeToArray<UInt32>(ref _bytes, (UInt32)hnid.dwValue);
                    break;
                case EpropertyType.PtypBoolean:
                    _bytes = new byte[1] { (byte)hnid.dwValue };
                    break;
                case EpropertyType.PtypInteger64:
                case EpropertyType.PtypFloating64:
                    _bytes = new byte[8];
                    break;
                case EpropertyType.PtypTime:
                    _bytes = ValueToArray(FM.defaulTime);
                    break;
            }

        }
        public Property(PCBTH pc, byte[] value)
        {
            id = pc.wPropId;
            type = pc.wPropType;
            hnid = pc.dwValueHnid;
            _bytes = value;
        }
        public Property(EpropertyId propId, EpropertyType propType)
        {
            id = propId;
            type = propType;
            hnid = new HNID();
            _bytes = null;
            switch (type)
            {
                case EpropertyType.PtypInteger16:
                    _bytes = new byte[2];
                    break;
                case EpropertyType.PtypInteger32:
                    _bytes = new byte[4];
                    break;
                case EpropertyType.PtypBoolean:
                    _bytes = new byte[1];
                    break;
                case EpropertyType.PtypInteger64:
                case EpropertyType.PtypFloating64:
                    _bytes = new byte[8];
                    hnid = new HNID(32);
                    break;
                case EpropertyType.PtypTime:
                    _bytes = ValueToArray(FM.defaulTime);
                    hnid = new HNID(32);
                    break;
            }
        }
        public Property(EpropertyId propId, EpropertyType propType, byte[] value)
        {
            id = propId;
            type = propType;
            hnid = new HNID(new HID(EnidType.HID, 1, 0));
            _bytes = value;
        }
        public Property(EpropertyId propId, EpropertyType propType, dynamic value)
        {
            id = propId;
            type = propType;
            hnid = new HNID(new HID(EnidType.HID, 1, 0));
            _bytes = ValueToArray(value);
            if (propType == EpropertyType.PtypString && value == "")
            {
                hnid = new HNID(0);
            }
        }

        private byte[] ValueToArray(dynamic value)
        {
            byte[] array = null;
            switch (type)
            {
                case EpropertyType.PtypInteger16:
                    array = new byte[2];
                    FM.TypeToArray<UInt16>(ref array, (UInt16)value);
                    hnid = new HNID((UInt32)value);
                    break;
                case EpropertyType.PtypInteger32:
                    array = new byte[4];
                    FM.TypeToArray<UInt32>(ref array, (UInt32)value);
                    hnid = new HNID((UInt32)value);
                    break;
                case EpropertyType.PtypInteger64:
                    array = new byte[8];
                    FM.TypeToArray<UInt64>(ref array, (UInt64)value);
                    break;
                case EpropertyType.PtypFloating64:
                    array = new byte[8];
                    FM.TypeToArray<Double>(ref array, (Double)value);
                    break;
                case EpropertyType.PtypMultipleInteger32:
                    /*
                    tbd
                    */
                    break;

                case EpropertyType.PtypBoolean:
                    hnid = new HNID(0);
                    if (value)
                    {
                        array = new byte[1] { 1 };
                        hnid = new HNID(1);
                    }
                    else
                    {
                        array = new byte[1] { 0 };
                        hnid = new HNID(0);
                    }
                    break;

                case EpropertyType.PtypBinary:
                    array = (byte[])value;
                    break;

                case EpropertyType.PtypString: // Unicode string
                    array = Encoding.Unicode.GetBytes(value);
                    break;

                case EpropertyType.PtypString8:  // Multipoint string in variable encoding
                    array = Encoding.UTF8.GetBytes(value);
                    break;

                case EpropertyType.PtypMultipleString: // Unicode strings
                    /*
                    tbd
                    */
                    break;

                case EpropertyType.PtypMultipleBinary:
                    /*
                    tbd
                    */
                    break;

                case EpropertyType.PtypTime:
                    array = new byte[8];
                    FM.TypeToArray<UInt64>(ref array, (UInt64)value);   // need to test it!!!!!!!!!
                    break;

                case EpropertyType.PtypGuid:
                    /*
                    tbd
                    */
                    break;

                case EpropertyType.PtypObject:
                    /*
                    tbd
                    */
                    break;

                default:
                    throw new Exception($"Unsupported property type {type}");
                    break;
            }
            return array;
        }

        public void setHNID(HNID Hnid)
        {
            hnid = Hnid;
        }
        public int dataSize { get { return (_bytes == null) ? 0 : _bytes.Length; } }
        public bool hasVariableSize => !hasFixedSize;
        public bool hasFixedSize
        {
            get
            {
                switch (type)
                {
                    case EpropertyType.PtypInteger16:
                    case EpropertyType.PtypInteger32:
                    case EpropertyType.PtypInteger64:
                    case EpropertyType.PtypFloating64:
                    case EpropertyType.PtypBoolean:
                    case EpropertyType.PtypTime:
                        return true;
                    default:
                        return false;
                }
            }
        }
    }
}
