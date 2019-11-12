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

        worksheet.write(0, 1, 'Apex Class', header_format)  
        worksheet.write(0, 2, 'Approval Process', header_format)  
        worksheet.write(0, 3, 'Custom Metadata', header_format)  
        worksheet.write(0, 4, 'Field', header_format)  
        worksheet.write(0, 5, 'Flow', header_format)  
        worksheet.write(0, 6, 'Global Value Set', header_format)  
        worksheet.write(0, 7, 'Layout', header_format)  
        worksheet.write(0, 8, 'Page', header_format)  
        worksheet.write(0, 9, 'Quick Action', header_format)  
        worksheet.write(0, 10, 'Report', header_format)  
        worksheet.write(0, 11, 'Standard Value Set', header_format)  
        worksheet.write(0, 12, 'Trigger', header_format)  
        worksheet.write(0, 13, 'Validation Rule', header_format)  
        worksheet.write(0, 14, 'Workflow', header_format)  
    end

    def write_lines(file, worksheet)
        row = 1
        col = 0

        file.readlines.each do |line|
            worksheet.write(row, 0, line)

            worksheet.write(row, 1, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/*.cls'))
            worksheet.write(row, 2, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.approvalProcess-meta.xml'))
            worksheet.write(row, 3, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.md-meta.xml'))
            worksheet.write(row, 4, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.field-meta.xml'))
            worksheet.write(row, 5, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.flow-meta.xml'))
            worksheet.write(row, 6, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.globalValueSet-meta.xml'))
            worksheet.write(row, 7, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.layout-meta.xml'))
            worksheet.write(row, 8, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.page-meta.xml'))
            worksheet.write(row, 9, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.quickAction-meta.xml'))
            worksheet.write(row, 10, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.report-meta.xml'))
            worksheet.write(row, 11, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.standardValueSet.xml'))
            worksheet.write(row, 12, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.Trigger'))
            worksheet.write(row, 13, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.validationRule-meta.xml'))
            worksheet.write(row, 14, dependency_search(line, '/Users/daniel.m/Projects/Polaris/**/**.workflow-meta.xml'))

            puts row + 'fields reviewed'
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