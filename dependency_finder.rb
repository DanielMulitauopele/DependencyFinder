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

    def run
        puts "Welcome to the DependencyFinder tool, created by Daniel Mulitauopele (11/11/19)." 
        puts "This tool is designed to track salesforce dependencies across different file types by running a search across your org source files." 
        puts "Please ensure that your source code is up to date before running the search."
        STDIN.getc

        puts "See below for the list of commands and their corresponding file types." 
        puts "In the event of missing file types, feel free to submit a pull request to the repository at: https://github.com/DanielMulitauopele/DependencyFinder." 
        puts "Once you know which file types you'd like to search, enter them below." 
        puts "If you'd like more than one dependency type, enter them separated by commas. Unseparated values may return faulty data. The commands are as follows:"
        STDIN.getc

        comm_count = @commands.keys.count - 1
        (1..comm_count).each do |key|
            puts "For #{@commands[key][:header]}, enter #{key}"
        end

        inputs = gets.chomp 
        formatted_inputs = inputs.split(',').map { |s| s.to_i }

        write_to_excel("opportunity_field_names.txt", formatted_inputs.unshift(0))
    end

    def write_to_excel(file, inputs)
        f = File.new(file, chomp: true)
        workbook = WriteXLSX.new('Opportunity_Field_Dependencies.xlsx')
        worksheet = workbook.add_worksheet
        
        write_lines(f, workbook, worksheet, inputs)

        workbook.close
        sleep 2
        puts 'Excel file created!'
    end

    def write_lines(file, workbook, worksheet, inputs)
        header_format = workbook.add_format(@font)
        inputs.each_with_index do |value, index|
            worksheet.write(0, index, @commands[value][:header], header_format) 
        end 

        row = 1
        file.readlines.each do |line|
            inputs.each_with_index do |value, index|
                if index == 0
                    worksheet.write(row, index, line.chomp)
                else
                    worksheet.write(row, index, dependency_search(line.chomp, @commands[value][:file_type]))
                end 
            end

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
                dependencies << file if line.downcase.include?(field.downcase)
            end
        end

        if dependencies.empty?
            return 'no results'
        else 
            return dependencies.uniq.join(', ')
        end
    end
end 