require 'ffi'
require 'byebug'
class CStrArray  < FFI::Struct
  layout :arr, :pointer,
         :s_size, :int
  
  def value
    self[:arr].read_array_of_pointer(self[:s_size]).map do |str_ptr| 
      str = str_ptr.read_string
      # str_ptr.free
      str
    end
  end
end

class CStrArray2 < FFI::Struct
  layout :arr, :pointer,
         :s_size, :int
  def value
    self[:arr].read_array_of_pointer(self[:s_size]).map do |p|
      strs = CStrArray.new(p).value
      # p.free
      strs
    end
  end
end


module XlsxExt
  extend FFI::Library
  ffi_lib "excel.lib"
  attach_function :printStr, [:string], :void
  
  attach_function :newFile, [:void], :uint32
  attach_function :openFile, [:string], :uint32
  attach_function :getSheetList, [:uint32], CStrArray.ptr
  attach_function :getSheetName, [:uint32, :int], :string

  attach_function :setCellValue, [:uint32, :string, :int, :int, :pointer, :int], :void
  attach_function :getCellValue, [:uint32, :string, :int, :int], :string

  attach_function :getRows, [:uint32, :string], CStrArray2.ptr
  
  attach_function :putRow, [:uint32, :string, :int, CStrArray.ptr], :void
  attach_function :putRows, [:uint32, :string, CStrArray2.ptr], :void

  attach_function :save, [:uint32], :void
  attach_function :saveAs, [:uint32, :string], :void

  attach_function :closeFile, [:uint32], :void
  
end

class Xlsx
  attr_accessor :id
  
  def self.new_file
    instance = Xlsx.new
    instance.id = XlsxExt.newFile()
    instance
  end

  def self.open_file(path)
    instance = Xlsx.new
    instance.id = XlsxExt.openFile(path)
    instance
  end

  def get_sheet_list
    ptr =  XlsxExt.getSheetList(self.id)
    ptr.value
  end

  def get_sheet_name(index)
    XlsxExt.getSheetName(self.id, index)
  end


  def set_cell_value(sheet_name, row, col, value)
    value = '' if value.nil?
    ptr = LibC.malloc(1)
    typ = if value.is_a?(Integer)
      ptr.write_int(value)
      1
    elsif value.is_a?(Float)
      ptr.write_float(value)
      2
    elsif value.is_a?(String)
      ptr.write_string(value)
      3
    end
    XlsxExt.setCellValue(self.id, sheet_name, row, col, ptr, typ)
    ptr.free
  end

  def get_cell_value(sheet_name, row, col)
    XlsxExt.getCellValue(self.id, sheet_name, row, col)
  end

  def get_rows(sheet_name)
    XlsxExt.getRows(self.id, sheet_name).value
  end

  def put_row(sheet_name, row_index, row)
    str = CStrArray.new
    str[:s_size] = row.size
    ptr = FFI::MemoryPointer.new(:pointer, row.size)
    pps = row.map{|x| FFI::MemoryPointer.from_string(x.to_s) }
    ptr.write_array_of_pointer(pps)
    str[:arr] = ptr
    XlsxExt.putRow(self.id, sheet_name, row_index,str)
    pps.each(&:free)
    ptr.free
    # str.free
  end

  def put_rows(sheet_name, rows)
    str2 = CStrArray2.new
    str2[:s_size] = rows.size
    ptr2 = FFI::MemoryPointer.new(:pointer, rows.size)
    todoRelease = [ptr2]
    
    ptr2_arr = rows.map do |row|
      str = CStrArray.new
      str[:s_size] = row.size

      ptr = FFI::MemoryPointer.new(:pointer, row.size)
      todoRelease << ptr
      sptrs = row.map{|s| FFI::MemoryPointer.from_string(s.to_s)}
      todoRelease += sptrs
      ptr.write_array_of_pointer(sptrs)
      str[:arr] = ptr
      str.pointer
    end
    todoRelease += ptr2_arr
    ptr2.write_array_of_pointer(ptr2_arr)
    str2[:arr] = ptr2
    XlsxExt.putRows(self.id, sheet_name, str2)
    todoRelease.each(&:free)
  end

  def save
    XlsxExt.save(self.id)
  end

  def save_as(path)
    XlsxExt.save(self.id,path)
  end

  def close
    XlsxExt.closeFile(self.id)
  end
end

module LibC
  extend FFI::Library
  ffi_lib FFI::Library::LIBC
  
  # memory allocators
  attach_function :malloc, [:size_t], :pointer
  attach_function :calloc, [:size_t], :pointer
  attach_function :valloc, [:size_t], :pointer
  attach_function :realloc, [:pointer, :size_t], :pointer
  attach_function :free, [:pointer], :void
  
  # memory movers
  attach_function :memcpy, [:pointer, :pointer, :size_t], :pointer
  attach_function :bcopy, [:pointer, :pointer, :size_t], :void
  
end # module LibC


# path = "/Users/wangdong/Desktop/aaa.xlsx"

# f = Xlsx.open_file(path)
# puts f.id
# puts f.get_sheet_name(4)
# a =  f.get_sheet_list
# puts 'aaaaa'
# puts a.join('|')

# f.close

# path2 = "/Users/wangdong/Desktop/wddd.xlsx"
# puts '1111'
# f2 = Xlsx.open_file(path2)
# puts 'abcd'
# puts f2.id
# puts f2.get_cell_value('Worksheet1', 2, 2)

# f2.set_cell_value("Worksheet1",13,4, 'abc')
# f2.set_cell_value("Worksheet1",13,5, 12)
# f2.set_cell_value("Worksheet1",13,6, 13.4234)
# f2.save
# tt = f2.get_rows('Worksheet1')
# puts tt[0][0]
# puts tt[0][1]
# f2.close



while true do
  path2 = "/Users/wangdong/Desktop/信息价网址页面.xlsx"
  puts path2
  f2 = Xlsx.open_file(path2)
  puts f2.get_rows('Sheet1')

  f2.put_row('Sheet1',5,[5,'b','c'])
  f2.put_rows('Sheet1',[[2,nil],[1,'b','c']])
  f2.save
  f2.close
end
