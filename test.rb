require 'ffi'
require 'byebug'
class CStrArray  < FFI::Struct
  layout :arr, :pointer,
         :str_size, :int
  
  def values
    self[:arr].read_array_of_pointer(self[:str_size]).map { |str_ptr| str_ptr.read_string }
  end

  def free
    LibC.free(self[:arr])
  end
end

class CStrArray2 < FFI::Struct
  layout :arr, :pointer,
         :row_sizes, :pointer,
         :rows,:int

  def values
    sizes = self[:row_sizes].read_array_of_int(self[:rows])
    LibC.free(self[:row_sizes])
    strs = self[:arr].read_array_of_pointer(sizes.sum).map do |str_ptr|
      str = str_ptr.read_string
      LibC.free(str_ptr)
      str
    end
    index = 0
    rows = []
    sizes.each do |str_size|
      rows.push(strs[index..index+str_size-1])
      index += str_size
    end
    rows
  end

  def free
    LibC.free(self[:arr])
    XlsxExt.freeStringArr2(self)
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

  attach_function :save, [:uint32], :void
  attach_function :saveAs, [:uint32, :string], :void

  attach_function :closeFile, [:uint32], :void
  attach_function :freeStringArr2, [:pointer], :void
  
end

class Xlsx
  attr_accessor :id
  
  def self.new_file
    instance = Xlsx.new
    instance.id = XlsxExt.newFile()
    return instance
  end

  def self.open_file(path)
    instance = Xlsx.new
    instance.id = XlsxExt.openFile(path)
    return instance
  end

  def get_sheet_list
    
    ptr =  XlsxExt.getSheetList(self.id)
    return ptr.values
  end

  def get_sheet_name(index)
    return XlsxExt.getSheetName(self.id, index)
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
    LibC.free(ptr)
  end

  def get_cell_value(sheet_name, row, col)
    XlsxExt.getCellValue(self.id, sheet_name, row, col)
  end

  def get_rows(sheet_name)
    XlsxExt.getRows(self.id, sheet_name).values
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
  f2.close
  GC.start
end
