require 'write_xlsx'
require 'find'
require 'pry'

class DependencyFinder
    def write_to_excel(file)
        f = File.new(file, chomp: true)
        workbook = WriteXLSX.new('Opportunity_Field_Dependencies.xlsx')
        worksheet = workbook.add_worksheet
        
        write_headers(workbook, worksheet)
        write_lines(f, worksheet)

        workbook.close
        puts 'Excel file created!'
    end

    def write_headers(workbook, worksheet)
        font = {
            :font  => 'Calibri',
            :size  => 14,
            :color => 'black',
            :bold  => 1
        }

        header_format = workbook.add_format(font)

        worksheet.write(0, 0, 'API Name', header_format)    
        worksheet.write(0, 1, 'Apex Class Dependencies', header_format)  
        worksheet.write(0, 2, 'Field Dependencies', header_format)  
        worksheet.write(0, 3, 'Layout Dependencies', header_format)  
        worksheet.write(0, 4, 'Trigger Dependencies', header_format)  
    end

    def write_lines(file, worksheet)
        row = 1
        col = 0

        file.readlines.each do |line|
            worksheet.write(row, 0, line)
            worksheet.write(row, 1, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/*.cls'))
            worksheet.write(row, 2, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.field-meta.xml'))
            worksheet.write(row, 3, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.layout-meta.xml'))
            worksheet.write(row, 4, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.Trigger'))

            puts row
            row += 1    
        end
    end

    def dependency_search(field, dependency_type)
        dependencies = []

        Dir.glob(dependency_type) do |file|
            File.readlines(file).each do |line|
                dependencies << file if line.include?(field.chomp)
            end
        end

        if dependencies.empty?
            return 'no results'
        else 
            return dependencies.uniq.join(', ')
        end
    end
end 