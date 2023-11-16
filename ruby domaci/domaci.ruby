require "google_drive"
$LOAD_PATH << 'C:/Users/aleks/.local/share/gem/ruby/3.2.0/gems/google_drive-3.0.7/lib'




session = GoogleDrive::Session.from_config("config.json")


ws = session.spreadsheet_by_key("1VkVi2kyLe8J-4Ek78OH_TB44gONx1eCJIaM6bW60TdM").worksheets[0]



GoogleDrive::Worksheet.class_eval do
  include Enumerable
  def each
    self.filtered_rows.each do |row|
      row.each do |cell|
        if self.check_cell(cell)
          yield cell
        end
      end
    end
  end
end

def add_method(c, m, &b)
  c.class_eval {
    define_method(m, &b)
  }
end

add_method(GoogleDrive::Worksheet, :matrix){
  matrix = []
  self.filtered_rows.each do |row|
    row_array = []
    row.each do |cell|
      if self.check_cell(cell)
        row_array << cell
      end
    end
    matrix << row_array
  end
  return matrix
}

GoogleDrive::Worksheet.class_eval do
  alias_method :original, :[]
  class Column
    def initialize(worksheet, column_name,&data)
      @worksheet = worksheet
      @column_name = column_name
    end
    
    def inspect
      if @worksheet.check_header(@column_name)
        index = @worksheet.get_index_of_column(@column_name)
        return @worksheet.filtered_rows.map[1..-1] { |row| row[index] unless row[index].empty?  }.compact
      end
    end
    
    def sum
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        column = @worksheet.filtered_rows[1..-1].map {|row| row[column_index]}.compact
        result = 0
        column.each do |cell|
          result += cell.to_i
        end
        return result
      else
        return nil
      end
    end

    def avg
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        column = @worksheet.filtered_rows[1..-1].map {|row| row[column_index]}.compact
        sum = 0
        size = 0
        column.each do |cell|
          cell_value = cell.to_i
          if @worksheet.check_cell(cell) && cell_value != 0
            size+=1
            sum += cell_value
          end
        end
        return sum/size
      else
        return nil
      end
    end
    
    def method_missing(cell)
      cell = cell[0..-1]
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        column = @worksheet.filtered_rows[1..-1].map {|row| row[column_index]}.compact
        if @worksheet.check_row(@worksheet.rows[column.index(cell)])
          return @worksheet.rows[column.index(cell)+1]
        else
          return nil
        end
      end
    end

    def map(&block)
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        result = []
        @worksheet.filtered_rows[1..-1].map do |row|
          if @worksheet.check_row(row) && check_cell(row[column_index])
            new_value = block.call(row[column_index])
            result << new_value
          end
        end
        return result
      else
        return nil
      end
    end
    
    def map!(&block)
      if check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        @worksheet.rows[1..-1].map do |row|
          if @worksheet.check_row(row)
            row_index = @worksheet.rows.index(row)
            if @worksheet.check_cell(@worksheet[row_index+1, column_index + 1])
              new_value = block.call(row[column_index])
              @worksheet[row_index + 1, column_index + 1] = new_value
              @worksheet.save
            end
          end
        end
      else
        return nil
      end
    end

    def select(&block)
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        result = []
        @worksheet.filtered_rows[1..-1].select do |row|
          if block.call(row[column_index]) && @worksheet.check_cell(row[column_index])
            #row[column_index].to_i ??
            result << row[column_index]
          end
        end
        return result
      else
        return nil
      end
    end
    #dodati destruktivni select koji upisuje prazan string kad se ne zadovolji uslov?

    def reduce(accumulator, &block)
      headers = @worksheet.rows[0]
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        result = accumulator
        @worksheet.filtered_rows[1..-1].each do |row|
          result = block.call(result, row[column_index].to_i)
        end
        return result
      else
        return nill
      end
    end


    def [](index)
      if @worksheet.check_header(@column_name)
        column_index = @worksheet.get_index_of_column(@column_name)
        column = @worksheet.rows.map {|row| row[column_index]}.compact
        column.each_with_index do |cell,index_in_loop|
          if index_in_loop >= index && @worksheet.check_cell(column[index_in_loop])
            return column[index_in_loop]
          end
        end
      else
        return nil
      end
    end

    def []=(index, value)
      if @worksheet.check_header(@column_name)
        row_index = get_row_index(index)
        column_index = @worksheet.get_index_of_column(@column_name)
        column = @worksheet.rows.map {|row| row[column_index]}.compact
        column.each_with_index do |cell,index_in_loop|
          if index_in_loop >= index && @worksheet.check_cell(column[index_in_loop])
            @worksheet[index_in_loop + 1, column_index + 1] = value
            @worksheet.save
            return
          end
        end
      else
        return nil
      end
    end
    
    def get_row_index(index)
      real_index = 0
      @worksheet.rows do |row|
        real_index+=1 if check_row(row)
        index-=1
        return real_index if index == 0
      end
      real_index
    end
  end
  
  def [](arg1, arg2 = nil)
    if arg2.nil?
      if check_header(arg1)
        return Column.new(self, arg1)
      else
        return nil
      end
    else
      original(arg1,arg2)
    end
  end
  
  def row(index)
     return self.rows[index].select {|cell| cell if self.check_cell(cell)}
  end
  
  def check_header(column_name)
    headers = self.rows[0]
    return headers.include?(column_name)
  end
  
  def check_row(row)
    row.each do |cell|
      if cell.casecmp?("total") || cell.casecmp?("subtotal")
        return false
      end
    end
    return true
  end
  
  def filtered_rows
    return self.rows.each.select {|row| row if self.check_row(row)}
  end
  
  def get_index_of_column(column_name)
    return self.rows[0].index(column_name)
  end
  
  def check_cell(cell)
    return !cell.empty?
  end
  
  def method_missing(column_name)
    column_name = column_name[0..-1]
    if self.check_header(column_name)
        return self[column_name] 
    end
  end
end

