require 'write_xlsx'
require 'find'
require 'pry'

class DependencyFinder
    def initialize
        @commands = {
            0 => {header: 'API Name'},
            1 => {header: 'Apex Class', file_type: '/Users/daniel.m/Projects/Polaris/**/*.cls'},
            2 => {header: 'Approval Process', file_type: '/Users/daniel.m/Projects/Polaris/**/**.approvalProcess-meta.xml'},
            3 => {header: 'Custom Metadata', file_type: '/Users/daniel.m/Projects/Polaris/**/**.md-meta.xml'},
            4 => {header: 'Field', file_type: '/Users/daniel.m/Projects/Polaris/**/**.field-meta.xml'},
            5 => {header: 'Flow', file_type: '/Users/daniel.m/Projects/Polaris/**/**.flow-meta.xml'},
            6 => {header: 'Global Value Set', file_type: '/Users/daniel.m/Projects/Polaris/**/**.globalValueSet-meta.xml'},
            7 => {header: 'Layout', file_type: '/Users/daniel.m/Projects/Polaris/**/**.layout-meta.xml'},
            8 => {header: 'Page', file_type: '/Users/daniel.m/Projects/Polaris/**/**.page-meta.xml'},
            9 => {header: 'Quick Action', file_type: '/Users/daniel.m/Projects/Polaris/**/**.quickAction-meta.xml'},
            10 => {header: 'Report', file_type: '/Users/daniel.m/Projects/Polaris/**/**.report-meta.xml'},
            11 => {header: 'Standard Value Set', file_type: '/Users/daniel.m/Projects/Polaris/**/**.standardValueSet.xml'},
            12 => {header: 'Trigger', file_type: '/Users/daniel.m/Projects/Polaris/**/**.Trigger'},
            13 => {header: 'Validation Rule', file_type: '/Users/daniel.m/Projects/Polaris/**/**.validationRule-meta.xml'},
            14 => {header: 'Workflow', file_type: '/Users/daniel.m/Projects/Polaris/**/**.workflow-meta.xml'}
        }

        @font = {
            :font  => 'Calibri',
            :size  => 14,
            :color => 'black',
            :bold  => 1
        }
    end

    def write_to_excel(file, inputs)
        f = File.new(file, chomp: true)
        workbook = WriteXLSX.new('Opportunity_Field_Dependencies.xlsx')
        worksheet = workbook.add_worksheet
        
        write_headers(workbook, worksheet, inputs)
        write_lines(f, worksheet, inputs)

        workbook.close
        sleep 2
        puts 'Excel file created!'
    end

    def write_headers(workbook, worksheet, inputs)
        header_format = workbook.add_format(@font)

        worksheet.write(0, 0, @commands[0][:header], header_format)

        worksheet.write(0, 1, @commands[1][:header], header_format)  
        worksheet.write(0, 2, @commands[2][:header], header_format)  
        worksheet.write(0, 3, @commands[3][:header], header_format)  
        worksheet.write(0, 4, @commands[4][:header], header_format)  
        worksheet.write(0, 5, @commands[5][:header], header_format)  
        worksheet.write(0, 6, @commands[6][:header], header_format)  
        worksheet.write(0, 7, @commands[7][:header], header_format)  
        worksheet.write(0, 8, @commands[8][:header], header_format)  
        worksheet.write(0, 9, @commands[9][:header], header_format)  
        worksheet.write(0, 10, @commands[10][:header], header_format)  
        worksheet.write(0, 11, @commands[11][:header], header_format)  
        worksheet.write(0, 12, @commands[12][:header], header_format)  
        worksheet.write(0, 13, @commands[13][:header], header_format)  
        worksheet.write(0, 14, @commands[14][:header], header_format)  
    end

    def write_lines(file, worksheet, inputs)
        row = 1
        col = 0

        file.readlines.each do |line|
            worksheet.write(row, 0, line)

            worksheet.write(row, 1, dependency_search(line, @commands[1][:file_type]))
            worksheet.write(row, 2, dependency_search(line, @commands[2][:file_type]))
            worksheet.write(row, 3, dependency_search(line, @commands[3][:file_type]))
            worksheet.write(row, 4, dependency_search(line, @commands[4][:file_type]))
            worksheet.write(row, 5, dependency_search(line, @commands[5][:file_type]))
            worksheet.write(row, 6, dependency_search(line, @commands[6][:file_type]))
            worksheet.write(row, 7, dependency_search(line, @commands[7][:file_type]))
            worksheet.write(row, 8, dependency_search(line, @commands[8][:file_type]))
            worksheet.write(row, 9, dependency_search(line, @commands[9][:file_type]))
            worksheet.write(row, 10, dependency_search(line, @commands[10][:file_type]))
            worksheet.write(row, 11, dependency_search(line, @commands[11][:file_type]))
            worksheet.write(row, 12, dependency_search(line, @commands[12][:file_type]))
            worksheet.write(row, 13, dependency_search(line, @commands[13][:file_type]))
            worksheet.write(row, 14, dependency_search(line, @commands[14][:file_type]))

            output = row.to_s + ' fields reviewed...' + "\r"
            print output
            $stdout.flush
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